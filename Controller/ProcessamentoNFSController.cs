using System;
using System.Collections.Generic;
using System.Linq;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Services;
using SAPbobsCOM;
using System.Configuration;
using System.Threading.Tasks;

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
            List<ResultadoProcessamento> resultados = new List<ResultadoProcessamento>();

            try
            {
                _grupoController.AtualizarStatusGrupoSimples(grupo.Code, StatusGrupo.EmProcessamento);

                for (int i = 0; i < linhas.Count; i++)
                {
                    var linha = linhas[i];
                    var resultado = ProcessarLinha(grupo.TipoDocumento, linha, grupo); // Chama o método síncrono
                    resultados.Add(resultado);

                    try
                    {
                        // Atualiza o status principal da linha no banco
                        AtualizarStatusLinha(linha, resultado);
                    }
                    catch (Exception ex)
                    {
                        resultado.Mensagem += $" | Erro ao salvar status principal: {ex.Message}";
                    }
                }

                bool sucessoSalvamento = RecalcularStatusGrupoEficiente(grupo.Code);
                if (!sucessoSalvamento)
                {
                    System.Diagnostics.Debug.WriteLine($"AVISO: Falha ao recalcular totais do grupo {grupo.Code}.");
                }
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
            var tempoInicio = DateTime.Now;
            List<ResultadoProcessamento> resultados = new List<ResultadoProcessamento>();

            try
            {
                var validacao = ValidarDadosSAP(linhas);
                if (!validacao.Valida && validacao.Erros.Count > 0)
                {
                    throw new Exception($"Erros de validação: {string.Join(", ", validacao.Erros.Take(3))}");
                }

                _grupoController.AtualizarStatusGrupoSimples(grupo.Code, StatusGrupo.EmProcessamento);

                for (int i = 0; i < linhas.Count; i++)
                {
                    var linha = linhas[i];
                    var resultado = ProcessarLinha(grupo.TipoDocumento, linha, grupo); // Chama o mesmo método síncrono
                    resultados.Add(resultado);

                    // Apenas atualiza o objeto em memória por enquanto
                    linha.Status = resultado.Sucesso ? StatusLinha.Sucesso : StatusLinha.Erro;
                    linha.DocNum = resultado.DocNum;
                    linha.DocEntry = resultado.DocEntry;
                    linha.MensagemErro = resultado.Mensagem;
                    linha.Reprocessar = !resultado.Sucesso;
                }

                // Atualiza o status de todas as linhas processadas no banco de uma vez
                AtualizarStatusLinhasEmLote(linhas, resultados);
                RecalcularStatusGrupoEficiente(grupo.Code);

                var tempoTotal = DateTime.Now - tempoInicio;
                System.Diagnostics.Debug.WriteLine($"TOTAL OTIMIZADO: {tempoTotal.TotalSeconds:F1}s para {linhas.Count} linhas");
            }
            catch (Exception ex)
            {
                try { _grupoController.AtualizarStatusGrupoSimples(grupo.Code, StatusGrupo.Erro); } catch { }
                throw new Exception($"Erro no processamento otimizado: {ex.Message}", ex);
            }

            return resultados;
        }


        #endregion

        #region Lógica de Processamento de Linha Única (Privado e Centralizado)

        /// <summary>
        /// Processa uma única linha de forma síncrona. Ponto central da lógica.
        /// </summary>
        private ResultadoProcessamento ProcessarLinha(string tipoDocumento, LinhaImportacao linha, GrupoLote grupo)
        {
            var resultado = new ResultadoProcessamento
            {
                CodigoLinha = linha.Code,
                NumeroLinha = linha.NumeroLinha
            };

            try
            {
                switch (tipoDocumento)
                {
                    case "NFS":
                        // Chama o método síncrono para criar a nota fiscal
                        return CriarNotaFiscalSaida(linha, grupo, resultado);

                    case "ENT":
                        // Chama o método síncrono para criar a entrega (que agora inclui a geração do PDF)
                        return CriarEntrega(linha, grupo, resultado);

                    case "NFE":
                        // Chama o método síncrono para criar a nota fiscal de entrada
                        return CriarNotaFiscalEntrada(linha, grupo, resultado);

                    default:
                        throw new InvalidOperationException($"Tipo de documento desconhecido ou não informado: '{tipoDocumento}'");
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
                resultado.Mensagem = $"Erro na linha {linha.NumeroLinha}: {ex.Message}";
                if (ex.InnerException != null)
                {
                    resultado.Mensagem += $" | {ex.InnerException.Message}";
                }
            }

            return resultado;
        }

        #endregion

        #region Métodos de Criação de Documentos (Privados)

        private ResultadoProcessamento CriarNotaFiscalSaida(LinhaImportacao linha, GrupoLote grupo, ResultadoProcessamento resultado)
        {
            // ... (A lógica de montar o invoiceRequest permanece a mesma)
            var invoiceRequest = new ServiceLayerInvoiceClient.InvoiceRequest { /* ... preencher dados ... */ };
            var invoiceResponse = _invoiceClient.CreateInvoice(invoiceRequest);

            resultado.Sucesso = true;
            resultado.DocEntry = invoiceResponse.DocEntry;
            resultado.DocNum = invoiceResponse.DocNum;
            resultado.Mensagem = $"NFS-e criada - DocNum: {invoiceResponse.DocNum}";

            // PDF não se aplica a este tipo de documento, então não fazemos nada.
            return resultado;
        }

        private ResultadoProcessamento CriarNotaFiscalEntrada(LinhaImportacao linha, GrupoLote grupo, ResultadoProcessamento resultado)
        {
            // ... (A lógica de montar o requestData permanece a mesma) ...
            var requestData = new ServiceLayerInvoiceClient.InvoiceRequest { /* ... preencher dados ... */ };
            var response = _invoiceClient.CreatePurchaseInvoice(requestData);

            resultado.Sucesso = true;
            resultado.DocEntry = response.DocEntry;
            resultado.DocNum = response.DocNum;
            resultado.Mensagem = $"NF de Entrada criada - DocNum: {response.DocNum}";

            // PDF não se aplica a este tipo de documento.
            return resultado;
        }

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

            // 2. Tenta gerar o PDF via API Gateway
            string pdfStatus = "N";
            string pdfMsg = "";
            try
            {
                string layoutCode = ConfigurationManager.AppSettings["DeliveryNote_LayoutCode"];
                string pastaDestino = grupo.CaminhoPDF;
                string nomeArquivo = $"ENT_{linha.CodigoCliente}_{response.DocNum}_{grupo.DataDocumento:yyyyMMdd}";

                // Chama a versão síncrona do serviço de PDF, que faz o bloqueio necessário
                var pdfResult = _pdfGenerationService.GerarPdfDeEntrega(
                    response.DocEntry,
                    pastaDestino,
                    nomeArquivo
                );

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
                    resultado.Mensagem += $" | PDF: Erro";
                    resultado.PdfGeradoComSucesso = false;
                    resultado.PdfMensagemErro = pdfMsg;
                }
            }
            catch (Exception pdfEx)
            {
                pdfStatus = "E";
                pdfMsg = pdfEx.Message;
                resultado.Mensagem += $" | PDF: Exceção";
                resultado.PdfGeradoComSucesso = false;
                resultado.PdfMensagemErro = pdfMsg;
            }
            finally
            {
                // 3. Atualiza o status do PDF no banco, independentemente do resultado
                AtualizarStatusPdfLinha(linha.Code, pdfStatus, pdfMsg);
                linha.PdfStatus = pdfStatus;
                linha.PdfMsg = pdfMsg;
            }

            return resultado;
        }

        #endregion

        #region Métodos de Atualização de Banco

        private void AtualizarStatusPdfLinha(string linhaCode, string status, string mensagem)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string msgSql = (mensagem ?? "").Replace("'", "''");
                if (msgSql.Length > 1000) msgSql = msgSql.Substring(0, 1000);

                string query = $@"UPDATE ""@IT_LINHA_LOTE""
                                SET ""U_PdfStatus"" = '{status}', ""U_PdfMsg"" = '{msgSql}'
                                WHERE ""Code"" = '{linhaCode}'";
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

        // Os métodos RecalcularStatusGrupoEficiente, AtualizarStatusLinha, AtualizarStatusLinhasEmLote,
        // e todos os métodos de Validação permanecem os mesmos que você já tem.
        // ... (Cole aqui os métodos existentes: RecalcularStatusGrupoEficiente, AtualizarStatusLinha,
        //      AtualizarStatusLinhasEmLote, e toda a região #region Métodos de Validação) ...

        #endregion

        // INCLUA AQUI O RESTANTE DO CÓDIGO QUE VOCÊ JÁ POSSUI (RecalcularStatus, AtualizarStatusLinha, etc.)
        // >>>>>>>>>>>>>>>>
        #region Método Centralizado de Recálculo de Status

        /// <summary>
        /// Recalcula e atualiza status do grupo usando query SQL eficiente
        /// </summary>
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

                if (oRecordset.RecordCount == 0)
                    return false;

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

        #endregion

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

        #region Métodos de Validação
        public ValidacaoImportacao ValidarDadosSAP(List<LinhaImportacao> linhas)
        {
            ValidacaoImportacao validacao = new ValidacaoImportacao();
            HashSet<string> clientesValidados = new HashSet<string>();
            HashSet<string> itensValidados = new HashSet<string>();
            foreach (var linha in linhas)
            {
                if (!string.IsNullOrEmpty(linha.CodigoCliente) && !clientesValidados.Contains(linha.CodigoCliente))
                {
                    if (!ValidarCliente(linha.CodigoCliente)) { validacao.AdicionarErro($"Cliente '{linha.CodigoCliente}' não encontrado no SAP"); }
                    clientesValidados.Add(linha.CodigoCliente);
                }
                if (!string.IsNullOrEmpty(linha.CodigoItem) && !itensValidados.Contains(linha.CodigoItem))
                {
                    if (!ValidarItem(linha.CodigoItem)) { validacao.AdicionarErro($"Item '{linha.CodigoItem}' não encontrado no SAP"); }
                    itensValidados.Add(linha.CodigoItem);
                }
                if (linha.Valor <= 0) { validacao.AdicionarErro($"Linha {linha.NumeroLinha}: Valor deve ser maior que zero"); }
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

        private bool ValidarUtilizacao(string usage)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery($@"SELECT ""ID"" FROM OUSG WHERE ""ID"" = '{usage}'");
                return !oRecordset.EoF;
            }
            catch { return false; }
            finally { if (oRecordset != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset); }
        }

        private bool ValidarCodigoImposto(string taxCode)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery($@"SELECT ""Code"" FROM OSTC WHERE ""Code"" = '{taxCode}'");
                return !oRecordset.EoF;
            }
            catch { return false; }
            finally { if (oRecordset != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset); }
        }

        private bool ValidarSequenciaNF(string seqCode)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery($@"SELECT ""SeqCode"" FROM NFN1 WHERE ""SeqCode"" = '{seqCode}' AND ""ObjectCode"" = '13'");
                return !oRecordset.EoF;
            }
            catch { return false; }
            finally { if (oRecordset != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset); }
        }

        private bool ValidarCondicaoPagamento(string groupNum)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery($@"SELECT ""GroupNum"" FROM OCTG WHERE ""GroupNum"" = '{groupNum}'");
                return !oRecordset.EoF;
            }
            catch { return false; }
            finally { if (oRecordset != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset); }
        }
        #endregion

        #region Métodos de Diagnóstico
        public bool CorrigirTotaisGrupo(string grupoCode)
        {
            try { return RecalcularStatusGrupoEficiente(grupoCode); }
            catch { return false; }
        }
        #endregion
        // <<<<<<<<<<<<<<<<
    }
}
