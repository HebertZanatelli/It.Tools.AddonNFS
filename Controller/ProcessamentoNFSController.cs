using System;
using System.Collections.Generic;
using System.Linq;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Services;
using SAPbobsCOM;
using System.Configuration;

namespace ItTech.Tool.AddonNFS.Controllers
{
    /// <summary>
    /// Controller para processamento de NFS-e usando Service Layer
    /// </summary>
    public class ProcessamentoNFSController
    {
        private readonly Company _company;
        private readonly ServiceLayerInvoiceClient _invoiceClient;
        private readonly GrupoLoteController _grupoController;
        private readonly PdfGenerationService _pdfGenerationService;


        public ProcessamentoNFSController(Company company, ServiceLayerInvoiceClient invoiceClient)
        {
            _company = company ?? throw new ArgumentNullException(nameof(company));
            _invoiceClient = invoiceClient ?? throw new ArgumentNullException(nameof(invoiceClient));
            _grupoController = new GrupoLoteController(company);
            _pdfGenerationService = new PdfGenerationService(company);
        }

        #region Métodos de Processamento em Lote (Públicos)

        /// <summary>
        /// Processa múltiplas linhas, atualizando o status no banco a cada linha.
        /// </summary>
        public List<ResultadoProcessamento> ProcessarLinhas(GrupoLote grupo, List<LinhaImportacao> linhas)
        {
            var resultados = new List<ResultadoProcessamento>();
            try
            {
                _grupoController.AtualizarStatusGrupoSimples(grupo.Code, StatusGrupo.EmProcessamento);

                foreach (var linha in linhas)
                {
                    var resultado = ProcessarLinha(grupo.TipoDocumento, linha, grupo);
                    resultados.Add(resultado);

                    // A atualização do status do PDF já é feita dentro de ProcessarLinha.
                    // Aqui atualizamos o status principal da linha.
                    AtualizarStatusLinha(linha, resultado);
                }

                RecalcularStatusGrupoEficiente(grupo.Code);
            }
            catch (Exception ex)
            {
                try { _grupoController.AtualizarStatusGrupoSimples(grupo.Code, StatusGrupo.Erro); } catch { }
                throw new Exception($"Erro no processamento: {ex.Message}", ex);
            }
            return resultados;
        }

        /// <summary>
        /// Processa múltiplas linhas e atualiza o status de todas no banco ao final (em lote).
        /// </summary>
        public List<ResultadoProcessamento> ProcessarLinhasOtimizado(GrupoLote grupo, List<LinhaImportacao> linhas)
        {
            var resultados = new List<ResultadoProcessamento>();
            try
            {
                var validacao = ValidarDadosSAP(linhas);
                if (!validacao.Valida && validacao.Erros.Any())
                {
                    throw new Exception($"Erros de validação: {string.Join(", ", validacao.Erros.Take(3))}");
                }

                _grupoController.AtualizarStatusGrupoSimples(grupo.Code, StatusGrupo.EmProcessamento);

                foreach (var linha in linhas)
                {
                    // >>> CORREÇÃO APLICADA: Chama o mesmo método centralizado para garantir consistência.
                    var resultado = ProcessarLinha(grupo.TipoDocumento, linha, grupo);
                    resultados.Add(resultado);

                    // Apenas atualiza o objeto em memória para o salvamento em lote posterior.
                    linha.Status = resultado.Sucesso ? StatusLinha.Sucesso : StatusLinha.Erro;
                    linha.DocNum = resultado.DocNum;
                    linha.DocEntry = resultado.DocEntry;
                    linha.MensagemErro = resultado.Mensagem;
                    linha.Reprocessar = !resultado.Sucesso;
                }

                AtualizarStatusLinhasEmLote(linhas, resultados);
                RecalcularStatusGrupoEficiente(grupo.Code);
            }
            catch (Exception ex)
            {
                try { _grupoController.AtualizarStatusGrupoSimples(grupo.Code, StatusGrupo.Erro); } catch { }
                throw new Exception($"Erro no processamento otimizado: {ex.Message}", ex);
            }
            return resultados;
        }

        #endregion

        #region Lógica de Processamento de Linha Única (Centralizada)

        /// <summary>
        /// Processa uma única linha de forma síncrona. Ponto central da lógica.
        /// </summary>
        private ResultadoProcessamento ProcessarLinha(string tipoDocumento, LinhaImportacao linha, GrupoLote grupo)
        {
            var resultado = new ResultadoProcessamento { CodigoLinha = linha.Code, NumeroLinha = linha.NumeroLinha };
            try
            {
                switch (tipoDocumento)
                {
                    case "NFS":
                        return CriarNotaFiscalSaida(linha, grupo, resultado);
                    case "ENT":
                        return CriarEntrega(linha, grupo, resultado);
                    case "NFE":
                        return CriarNotaFiscalEntrada(linha, grupo, resultado);
                    default:
                        throw new InvalidOperationException($"Tipo de documento desconhecido: '{tipoDocumento}'");
                }
            }
            catch (ServiceLayerInvoiceClient.ServiceLayerException ex)
            {
                resultado.Sucesso = false;
                resultado.Mensagem = $"[{ex.ServiceLayerCode}] {ex.Message}";
            }
            catch (Exception ex)
            {
                resultado.Sucesso = false;
                resultado.Mensagem = $"Erro: {ex.Message}";
                if (ex.InnerException != null)
                {
                    resultado.Mensagem += $" | Detalhe: {ex.InnerException.Message}";
                }
            }
            return resultado;
        }

        #endregion

        #region Métodos de Criação de Documentos

        private ResultadoProcessamento CriarEntrega(LinhaImportacao linha, GrupoLote grupo, ResultadoProcessamento resultado)
        {
            // 1. Cria a Entrega via Service Layer
            var requestData = new ServiceLayerInvoiceClient.InvoiceRequest
            {
                CardCode = linha.CodigoCliente,
                DocDate = grupo.DataDocumento,
                BPL_IDAssignedToInvoice = linha.Filial,
                OpeningRemarks = linha.ObservacaoNF ?? "Gerado via Add-on de Lote",
                PaymentGroupCode = Convert.ToInt32(linha.CondicaoPagamento),
                SequenceCode = Convert.ToInt32(linha.CodSeq),
                DocumentLines = new List<ServiceLayerInvoiceClient.InvoiceDocumentLine>
                {
                    new ServiceLayerInvoiceClient.InvoiceDocumentLine
                    {
                        ItemCode = linha.CodigoItem,
                        Quantity = 1,
                        UnitPrice = linha.Valor,
                        TaxCode = linha.CodigoImposto
                    }
                }
            };
            var response = _invoiceClient.CreateDeliveryNote(requestData);

            resultado.Sucesso = true;
            resultado.DocEntry = response.DocEntry;
            resultado.DocNum = response.DocNum;
            resultado.Mensagem = $"Entrega criada - DocNum: {response.DocNum}";

            // 2. Tenta gerar o PDF
            string pdfStatus = "N";
            string pdfMsg = "";
            try
            {
                if (string.IsNullOrWhiteSpace(grupo.CaminhoPDF))
                {
                    throw new Exception("Caminho para salvar PDF não configurado no grupo.");
                }

                string nomeArquivo = $"ENT_{linha.CodigoCliente}_{response.DocNum}_{grupo.DataDocumento:yyyyMMdd}";

                // >>> CHAMADA SÍNCRONA (BLOQUEANTE) AO SERVIÇO DE PDF
                var pdfResult = _pdfGenerationService.GerarPdfDeEntrega(response.DocEntry, grupo.CaminhoPDF, nomeArquivo);

                if (pdfResult.Success)
                {
                    pdfStatus = "S";
                    resultado.Mensagem += " | PDF: Sucesso";
                    resultado.PdfGeradoComSucesso = true;
                    resultado.PdfCaminhoCompleto = pdfResult.Message;
                }
                else
                {
                    pdfStatus = "E";
                    pdfMsg = pdfResult.Message;
                    resultado.Mensagem += " | PDF: Erro";
                    resultado.PdfGeradoComSucesso = false;
                    resultado.PdfMensagemErro = pdfMsg;
                }
            }
            catch (Exception pdfEx)
            {
                pdfStatus = "E";
                pdfMsg = pdfEx.Message;
                resultado.Mensagem += " | PDF: Exceção";
                resultado.PdfGeradoComSucesso = false;
                resultado.PdfMensagemErro = pdfMsg;
            }
            finally
            {
                // 3. Atualiza o status do PDF no banco, aconteça o que acontecer.
                AtualizarStatusPdfLinha(linha.Code, pdfStatus, pdfMsg);
                linha.PdfStatus = pdfStatus;
                linha.PdfMsg = pdfMsg;
            }
            return resultado;
        }

        private ResultadoProcessamento CriarNotaFiscalSaida(LinhaImportacao linha, GrupoLote grupo, ResultadoProcessamento resultado)
        {
            var invoiceRequest = new ServiceLayerInvoiceClient.InvoiceRequest
            {
                CardCode = linha.CodigoCliente,
                DocDate = grupo.DataDocumento,
                BPL_IDAssignedToInvoice = linha.Filial,
                U_SKILL_TipTrib = linha.TipoTributacao,
                OpeningRemarks = linha.ObservacaoNF ?? "BANCO XXXX",
                PaymentGroupCode = Convert.ToInt32(linha.CondicaoPagamento),
                SequenceCode = Convert.ToInt32(linha.CodSeq),
                DocumentLines = new List<ServiceLayerInvoiceClient.InvoiceDocumentLine>
                {
                    new ServiceLayerInvoiceClient.InvoiceDocumentLine
                    {
                        ItemCode = linha.CodigoItem,
                        Quantity = 1,
                        UnitPrice = linha.Valor,
                        TaxCode = linha.CodigoImposto,
                        Usage = Convert.ToInt32(linha.Utilizacao)
                    }
                }
            };

            // Lógica da Regra do Município
            var modelSeqCode = ObterModelPeloSeqCode(int.Parse(linha.CodSeq));
            var cnpjFilial = ObterCNPJ(int.Parse(linha.Filial));
            string cnpjRegraSP = ConfigurationManager.AppSettings["CNPJRegraSP"];
            string cnpjRegraMG = ConfigurationManager.AppSettings["CNPJRegraMG"];

            if (modelSeqCode == "46" && cnpjFilial == cnpjRegraSP)
            {
                invoiceRequest.TaxExtension = new ServiceLayerInvoiceClient.InvoiceTaxExtension { State = "SP", County = "5215" };
            }

            if (modelSeqCode == "46" && cnpjFilial == cnpjRegraMG)
            {
                invoiceRequest.TaxExtension = new ServiceLayerInvoiceClient.InvoiceTaxExtension { State = "MG", County = "1410" };
            }

            var invoiceResponse = _invoiceClient.CreateInvoice(invoiceRequest);

            resultado.Sucesso = true;
            resultado.DocEntry = invoiceResponse.DocEntry;
            resultado.DocNum = invoiceResponse.DocNum;
            resultado.Mensagem = $"NFS-e criada - DocNum: {invoiceResponse.DocNum}";

            return resultado;
        }

        private ResultadoProcessamento CriarNotaFiscalEntrada(LinhaImportacao linha, GrupoLote grupo, ResultadoProcessamento resultado)
        {
            var requestData = new ServiceLayerInvoiceClient.InvoiceRequest
            {
                CardCode = linha.CodigoCliente,
                DocDate = grupo.DataDocumento,
                BPL_IDAssignedToInvoice = linha.Filial,
                OpeningRemarks = linha.ObservacaoNF ?? "Gerado via Add-on de Lote",
                PaymentGroupCode = Convert.ToInt32(linha.CondicaoPagamento),
                SequenceCode = Convert.ToInt32(linha.CodSeq),
                DocumentLines = new List<ServiceLayerInvoiceClient.InvoiceDocumentLine>
                {
                    new ServiceLayerInvoiceClient.InvoiceDocumentLine
                    {
                        ItemCode = linha.CodigoItem,
                        Quantity = 1,
                        UnitPrice = linha.Valor,
                        TaxCode = linha.CodigoImposto
                    }
                }
            };
            var response = _invoiceClient.CreatePurchaseInvoice(requestData);

            resultado.Sucesso = true;
            resultado.DocEntry = response.DocEntry;
            resultado.DocNum = response.DocNum;
            resultado.Mensagem = $"NF de Entrada criada - DocNum: {response.DocNum}";

            return resultado;
        }

        #endregion

        #region Métodos de Atualização, Validação e Diagnóstico (Completos)

        private void AtualizarStatusPdfLinha(string linhaCode, string status, string mensagem)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string msgSql = (mensagem ?? "").Replace("'", "''");
                if (msgSql.Length > 1000) msgSql = msgSql.Substring(0, 1000);

                string query = $@"UPDATE ""@IT_LINHA_LOTE"" SET ""U_PdfStatus"" = '{status}', ""U_PdfMsg"" = '{msgSql}' WHERE ""Code"" = '{linhaCode}'";
                rs.DoQuery(query);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao atualizar status do PDF para a linha {linhaCode}: {ex.Message}");
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        private bool RecalcularStatusGrupoEficiente(string grupoCode)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string queryCounts = $@"
                    SELECT 
                        COUNT(*) as Total,
                        COUNT(CASE WHEN ""U_Status"" = 'S' THEN 1 END) as Sucessos,
                        COUNT(CASE WHEN ""U_Status"" = 'E' THEN 1 END) as Erros,
                        COUNT(CASE WHEN ""U_Status"" = 'P' OR ""U_Status"" IS NULL THEN 1 END) as Pendentes
                    FROM ""@IT_LINHA_LOTE""
                    WHERE ""U_GrupoCode"" = '{grupoCode}'";

                oRecordset.DoQuery(queryCounts);

                if (oRecordset.RecordCount == 0) return false;

                int totalLinhas = Convert.ToInt32(oRecordset.Fields.Item("Total").Value ?? 0);
                int sucessos = Convert.ToInt32(oRecordset.Fields.Item("Sucessos").Value ?? 0);
                int erros = Convert.ToInt32(oRecordset.Fields.Item("Erros").Value ?? 0);
                int pendentes = Convert.ToInt32(oRecordset.Fields.Item("Pendentes").Value ?? 0);

                StatusGrupo statusFinal;
                if (totalLinhas == 0) { statusFinal = StatusGrupo.Erro; }
                else if (pendentes > 0) { statusFinal = StatusGrupo.ProcessadoParcial; }
                else if (erros > 0) { statusFinal = StatusGrupo.ProcessadoParcial; }
                else if (sucessos == totalLinhas) { statusFinal = StatusGrupo.ProcessadoCompleto; }
                else { statusFinal = StatusGrupo.Erro; }

                return _grupoController.AtualizarStatusGrupoComTotais(grupoCode, statusFinal, sucessos, erros, pendentes);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao recalcular status: {ex.Message}");
                return false;
            }
            finally
            {
                if (oRecordset != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        private void AtualizarStatusLinha(LinhaImportacao linha, ResultadoProcessamento resultado)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string status = resultado.Sucesso ? "S" : "E";
                string msgErro = (resultado.Mensagem ?? "").Replace("'", "''");
                if (msgErro.Length > 250) msgErro = msgErro.Substring(0, 250);

                string dataProcessamento = DateTime.Now.ToString("yyyy-MM-dd");
                int horaProcessamento = DateTime.Now.Hour * 60 + DateTime.Now.Minute;

                string query = $@"
                    UPDATE ""@IT_LINHA_LOTE""
                    SET ""U_Status"" = '{status}',
                        ""U_DocEntry"" = {(resultado.DocEntry.HasValue ? resultado.DocEntry.Value.ToString() : "NULL")},
                        ""U_DocNum"" = {(resultado.DocNum.HasValue ? resultado.DocNum.Value.ToString() : "NULL")},
                        ""U_MsgErro"" = '{msgErro}',
                        ""U_DataProc"" = '{dataProcessamento}',
                        ""U_HoraProc"" = {horaProcessamento},
                        ""U_Reprocessar"" = '{(resultado.Sucesso ? "N" : "Y")}'
                    WHERE ""Code"" = '{linha.Code}'";
                oRecordset.DoQuery(query);

                linha.Status = resultado.Sucesso ? StatusLinha.Sucesso : StatusLinha.Erro;
                linha.DocNum = resultado.DocNum;
                linha.MensagemErro = resultado.Mensagem;
                linha.Reprocessar = !resultado.Sucesso;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao atualizar status da linha {linha.NumeroLinha}: {ex.Message}", ex);
            }
            finally
            {
                if (oRecordset != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        private void AtualizarStatusLinhasEmLote(List<LinhaImportacao> linhas, List<ResultadoProcessamento> resultados)
        {
            if (linhas.Count != resultados.Count) return;
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                var comandos = new List<string>();
                for (int i = 0; i < linhas.Count; i++)
                {
                    var linha = linhas[i];
                    var resultado = resultados[i];
                    string status = resultado.Sucesso ? "S" : "E";
                    string msgErro = (resultado.Mensagem ?? "").Replace("'", "''");
                    if (msgErro.Length > 250) msgErro = msgErro.Substring(0, 250);
                    string dataProcessamento = DateTime.Now.ToString("yyyy-MM-dd");
                    int horaProcessamento = DateTime.Now.Hour * 60 + DateTime.Now.Minute;
                    string comando = $@"
                        UPDATE ""@IT_LINHA_LOTE""
                        SET ""U_Status"" = '{status}',
                            ""U_DocEntry"" = {(resultado.DocEntry.HasValue ? resultado.DocEntry.Value.ToString() : "NULL")},
                            ""U_DocNum"" = {(resultado.DocNum.HasValue ? resultado.DocNum.Value.ToString() : "NULL")},
                            ""U_MsgErro"" = '{msgErro}',
                            ""U_DataProc"" = '{dataProcessamento}',
                            ""U_HoraProc"" = {horaProcessamento},
                            ""U_Reprocessar"" = '{(resultado.Sucesso ? "N" : "Y")}'
                        WHERE ""Code"" = '{linha.Code}';";
                    comandos.Add(comando);
                }
                string queryCompleta = string.Join(" ", comandos);
                oRecordset.DoQuery(queryCompleta);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Erro no lote, tentando individual: {ex.Message}");
                for (int i = 0; i < linhas.Count; i++)
                {
                    try { AtualizarStatusLinha(linhas[i], resultados[i]); }
                    catch (Exception exIndividual) { System.Diagnostics.Debug.WriteLine($"⚠️ Erro linha {linhas[i].NumeroLinha}: {exIndividual.Message}"); }
                }
            }
            finally
            {
                if (oRecordset != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        public ValidacaoImportacao ValidarDadosSAP(List<LinhaImportacao> linhas)
        {
            ValidacaoImportacao validacao = new ValidacaoImportacao();
            HashSet<string> clientesValidados = new HashSet<string>();
            HashSet<string> itensValidados = new HashSet<string>();
            foreach (var linha in linhas)
            {
                if (!string.IsNullOrEmpty(linha.CodigoCliente) && !clientesValidados.Contains(linha.CodigoCliente))
                {
                    if (!ValidarCliente(linha.CodigoCliente)) { validacao.AdicionarErro($"Cliente '{linha.CodigoCliente}' não encontrado."); }
                    clientesValidados.Add(linha.CodigoCliente);
                }
                if (!string.IsNullOrEmpty(linha.CodigoItem) && !itensValidados.Contains(linha.CodigoItem))
                {
                    if (!ValidarItem(linha.CodigoItem)) { validacao.AdicionarErro($"Item '{linha.CodigoItem}' não encontrado."); }
                    itensValidados.Add(linha.CodigoItem);
                }
                if (linha.Valor <= 0) { validacao.AdicionarErro($"Linha {linha.NumeroLinha}: Valor deve ser maior que zero."); }
            }
            return validacao;
        }

        private bool ValidarCliente(string cardCode)
        {
            BusinessPartners oBP = null;
            try
            {
                oBP = (BusinessPartners)_company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                return oBP.GetByKey(cardCode);
            }
            catch { return false; }
            finally { if (oBP != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP); }
        }

        private bool ValidarItem(string itemCode)
        {
            Items oItem = null;
            try
            {
                oItem = (Items)_company.GetBusinessObject(BoObjectTypes.oItems);
                return oItem.GetByKey(itemCode);
            }
            catch { return false; }
            finally { if (oItem != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem); }
        }

        private string ObterModelPeloSeqCode(int seqCode)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT \"Model\" FROM NFN1 WHERE \"SeqCode\" = {seqCode}");
                if (rs.RecordCount > 0) return rs.Fields.Item("Model").Value.ToString();
                return string.Empty;
            }
            catch { return string.Empty; }
            finally { if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }

        private string ObterCNPJ(int bplId)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT \"TaxIdNum\" FROM OBPL WHERE \"BPLId\" = {bplId}");
                if (rs.RecordCount > 0) return rs.Fields.Item("TaxIdNum").Value.ToString();
                return string.Empty;
            }
            catch { return string.Empty; }
            finally { if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }

        public bool CorrigirTotaisGrupo(string grupoCode)
        {
            try { return RecalcularStatusGrupoEficiente(grupoCode); }
            catch { return false; }
        }
        #endregion
    }
}

