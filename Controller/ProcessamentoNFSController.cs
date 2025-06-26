using System;
using System.Collections.Generic;
using System.Linq;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Services;
using SAPbobsCOM;

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

        public ProcessamentoNFSController(Company company, ServiceLayerInvoiceClient invoiceClient)
        {
            _company = company ?? throw new ArgumentNullException(nameof(company));
            _invoiceClient = invoiceClient ?? throw new ArgumentNullException(nameof(invoiceClient));
            _grupoController = new GrupoLoteController(company);
        }

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

                // Query eficiente - apenas agregação, sem carregar objetos
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

                // Extrair contadores
                int totalLinhas = Convert.ToInt32(oRecordset.Fields.Item("Total").Value ?? 0);
                int sucessos = Convert.ToInt32(oRecordset.Fields.Item("Sucessos").Value ?? 0);
                int erros = Convert.ToInt32(oRecordset.Fields.Item("Erros").Value ?? 0);
                int pendentes = Convert.ToInt32(oRecordset.Fields.Item("Pendentes").Value ?? 0);

                // Determinar status com lógica correta (Pendentes têm prioridade!)
                StatusGrupo statusFinal;
                if (totalLinhas == 0)
                {
                    statusFinal = StatusGrupo.Erro;  // Grupo vazio?
                }
                else if (pendentes > 0)
                {
                    statusFinal = StatusGrupo.ProcessadoParcial;  // Prioridade 1
                }
                else if (erros > 0)
                {
                    statusFinal = StatusGrupo.ProcessadoParcial;  // Prioridade 2
                }
                else if (sucessos == totalLinhas)
                {
                    statusFinal = StatusGrupo.ProcessadoCompleto;  // Prioridade 3
                }
                else
                {
                    statusFinal = StatusGrupo.Erro;  // Fallback
                }

                // Log para debug
                System.Diagnostics.Debug.WriteLine(
                    $"RecalcularStatus: Grupo={grupoCode}, Total={totalLinhas}, " +
                    $"Sucessos={sucessos}, Erros={erros}, Pendentes={pendentes}, " +
                    $"Status={statusFinal}");

                // Atualizar usando método existente
                // IMPORTANTE: passar sucessos como "linhasProcessadas" (conforme semântica atual)
                return _grupoController.AtualizarStatusGrupoComTotais(
                    grupoCode,
                    statusFinal,
                    sucessos,    // ← "LinhasProcessadas" = sucessos (semântica atual)
                    erros,
                    pendentes
                );
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao recalcular status: {ex.Message}");
                return false;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        #endregion

        #region Métodos de Processamento

        /// <summary>
        /// Processa múltiplas linhas com otimizações de performance
        /// </summary>
        public List<ResultadoProcessamento> ProcessarLinhasOtimizado(string grupoCode, List<LinhaImportacao> linhas, DateTime dataLancamento, DateTime dataDocumento)
        {
            var tempoInicio = DateTime.Now;
            List<ResultadoProcessamento> resultados = new List<ResultadoProcessamento>();

            try
            {
                // Validação em lote
                var validacao = ValidarDadosSAP(linhas);
                if (!validacao.Valida && validacao.Erros.Count > 0)
                {
                    throw new Exception($"Erros de validação: {string.Join(", ", validacao.Erros.Take(3))}");
                }

                // Atualizar status para processamento
                _grupoController.AtualizarStatusGrupoSimples(grupoCode, StatusGrupo.EmProcessamento);

                // Processar documentos via Service Layer
                for (int i = 0; i < linhas.Count; i++)
                {
                    var linha = linhas[i];
                    var resultado = ProcessarLinha(linha, dataLancamento, dataDocumento);
                    resultados.Add(resultado);

                    // Atualizar objeto em memória
                    linha.Status = resultado.Sucesso ? StatusLinha.Sucesso : StatusLinha.Erro;
                    linha.DocNum = resultado.DocNum;
                    linha.DocEntry = resultado.DocEntry;
                    linha.MensagemErro = resultado.Mensagem;
                    linha.Reprocessar = !resultado.Sucesso;
                }

                // Atualizar banco em lote
                AtualizarStatusLinhasEmLote(linhas, resultados);

                // ✅ CORREÇÃO: Usar método centralizado
                RecalcularStatusGrupoEficiente(grupoCode);

                var tempoTotal = DateTime.Now - tempoInicio;
                System.Diagnostics.Debug.WriteLine($"🚀 TOTAL OTIMIZADO: {tempoTotal.TotalSeconds:F1}s para {linhas.Count} linhas");
            }
            catch (Exception ex)
            {
                try
                {
                    _grupoController.AtualizarStatusGrupoSimples(grupoCode, StatusGrupo.Erro);
                }
                catch { }

                throw new Exception($"Erro no processamento otimizado: {ex.Message}", ex);
            }

            return resultados;
        }

        /// <summary>
        /// Processa as linhas selecionadas criando as NFS-e
        /// </summary>
        public List<ResultadoProcessamento> ProcessarLinhas(string grupoCode, List<LinhaImportacao> linhas, DateTime dataLancamento, DateTime dataDocumento)
        {
            List<ResultadoProcessamento> resultados = new List<ResultadoProcessamento>();

            try
            {
                // Atualizar status para processamento
                _grupoController.AtualizarStatusGrupoSimples(grupoCode, StatusGrupo.EmProcessamento);

                // Processar cada linha
                for (int i = 0; i < linhas.Count; i++)
                {
                    var linha = linhas[i];
                    var resultado = ProcessarLinha(linha, dataLancamento, dataDocumento);
                    resultados.Add(resultado);

                    // Atualizar linha individual no banco
                    try
                    {
                        AtualizarStatusLinha(linha, resultado);
                    }
                    catch (Exception ex)
                    {
                        resultado.Mensagem += $" | Erro ao salvar: {ex.Message}";
                    }
                }

                // ✅ CORREÇÃO: Usar método centralizado ao invés de cálculo manual
                bool sucessoSalvamento = RecalcularStatusGrupoEficiente(grupoCode);

                if (!sucessoSalvamento)
                {
                    throw new Exception("Falha ao salvar totais do grupo no banco de dados");
                }
            }
            catch (Exception ex)
            {
                try
                {
                    _grupoController.AtualizarStatusGrupoSimples(grupoCode, StatusGrupo.Erro);
                }
                catch { }

                throw new Exception($"Erro no processamento: {ex.Message}", ex);
            }

            return resultados;
        }

        #endregion

        #region Métodos Privados

        /// <summary>
        /// Processa uma única linha - USANDO SERVICE LAYER
        /// </summary>
        private ResultadoProcessamento ProcessarLinha(LinhaImportacao linha, DateTime dataLancamento, DateTime dataDocumento)
        {
            ResultadoProcessamento resultado = new ResultadoProcessamento
            {
                CodigoLinha = linha.Code,
                NumeroLinha = linha.NumeroLinha,
                DataProcessamento = DateTime.Now
            };

            try
            {
                var invoiceRequest = new ServiceLayerInvoiceClient.InvoiceRequest
                {
                    CardCode = linha.CodigoCliente,
                    DocDate = dataDocumento,
                    DocDueDate = dataLancamento
                };

                // Filial
                if (!string.IsNullOrEmpty(linha.Filial))
                {
                    invoiceRequest.BPL_IDAssignedToInvoice = linha.Filial;
                }

                if(!string.IsNullOrEmpty(linha.TipoTributacao))
                {
                    invoiceRequest.U_SKILL_TipTrib = linha.TipoTributacao;
                }
                // Observações
                if (!string.IsNullOrEmpty(linha.ObservacaoNF))
                {
                    invoiceRequest.OpeningRemarks = linha.ObservacaoNF;
                }
                else
                {
                    invoiceRequest.OpeningRemarks = "BANCO XXXX";
                }

                // Condição de pagamento
                if (!string.IsNullOrEmpty(linha.CondicaoPagamento))
                {
                    try
                    {
                        invoiceRequest.PaymentGroupCode = Convert.ToInt32(linha.CondicaoPagamento);
                    }
                    catch
                    {
                        invoiceRequest.PaymentGroupCode = -1;
                    }
                }

                // Sequência do documento
                if (!string.IsNullOrEmpty(linha.CodSeq))
                {
                    try
                    {
                        invoiceRequest.SequenceCode = Convert.ToInt32(linha.CodSeq);
                    }
                    catch { }
                }

                // Linha do documento
                var documentLine = new ServiceLayerInvoiceClient.InvoiceDocumentLine
                {
                    ItemCode = linha.CodigoItem,
                    Quantity = 1,
                    UnitPrice = linha.Valor
                };

                // Código de imposto
                if (!string.IsNullOrEmpty(linha.CodigoImposto))
                {
                    documentLine.TaxCode = linha.CodigoImposto;
                }

                // Utilização
                if (!string.IsNullOrEmpty(linha.Utilizacao))
                {
                    try
                    {
                        documentLine.Usage = Convert.ToInt32(linha.Utilizacao);
                    }
                    catch { }
                }

                invoiceRequest.DocumentLines.Add(documentLine);

                // Criar via Service Layer
                var invoiceResponse = _invoiceClient.CreateInvoice(invoiceRequest);

                resultado.Sucesso = true;
                resultado.DocEntry = invoiceResponse.DocEntry;
                resultado.DocNum = invoiceResponse.DocNum;
                resultado.Mensagem = $"NFS-e criada - DocNum: {invoiceResponse.DocNum}";
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

        /// <summary>
        /// Atualiza o status de uma linha após o processamento
        /// </summary>
        private void AtualizarStatusLinha(LinhaImportacao linha, ResultadoProcessamento resultado)
        {
            Recordset oRecordset = null;

            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string status = resultado.Sucesso ? "S" : "E";
                string msgErro = "";
                if (!string.IsNullOrEmpty(resultado.Mensagem))
                {
                    msgErro = resultado.Mensagem.Replace("'", "''");
                    if (msgErro.Length > 250)
                    {
                        msgErro = msgErro.Substring(0, 250);
                    }
                }

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

                // Atualizar objeto em memória
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
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// Atualiza status de múltiplas linhas em uma única transação
        /// </summary>
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
                    string msgErro = "";
                    if (!string.IsNullOrEmpty(resultado.Mensagem))
                    {
                        msgErro = resultado.Mensagem.Replace("'", "''");
                        if (msgErro.Length > 250)
                        {
                            msgErro = msgErro.Substring(0, 250);
                        }
                    }

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

                // Executar comandos em lote
                string queryCompleta = string.Join(" ", comandos);
                oRecordset.DoQuery(queryCompleta);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ Erro no lote, tentando individual: {ex.Message}");

                // Fallback: atualizar uma por uma
                for (int i = 0; i < linhas.Count; i++)
                {
                    try
                    {
                        AtualizarStatusLinha(linhas[i], resultados[i]);
                    }
                    catch (Exception exIndividual)
                    {
                        System.Diagnostics.Debug.WriteLine($"⚠️ Erro linha {linhas[i].NumeroLinha}: {exIndividual.Message}");
                    }
                }
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        #endregion

        #region Métodos de Validação

        /// <summary>
        /// Valida se os dados necessários estão cadastrados no SAP
        /// </summary>
        public ValidacaoImportacao ValidarDadosSAP(List<LinhaImportacao> linhas)
        {
            ValidacaoImportacao validacao = new ValidacaoImportacao();
            HashSet<string> clientesValidados = new HashSet<string>();
            HashSet<string> itensValidados = new HashSet<string>();
            HashSet<string> utilizacoesValidadas = new HashSet<string>();
            HashSet<string> impostosValidados = new HashSet<string>();
            HashSet<string> sequenciasValidadas = new HashSet<string>();
            HashSet<string> condicoesValidadas = new HashSet<string>();

            foreach (var linha in linhas)
            {
                // Validar Cliente
                if (!string.IsNullOrEmpty(linha.CodigoCliente) && !clientesValidados.Contains(linha.CodigoCliente))
                {
                    if (!ValidarCliente(linha.CodigoCliente))
                    {
                        validacao.AdicionarErro($"Cliente '{linha.CodigoCliente}' não encontrado no SAP");
                    }
                    clientesValidados.Add(linha.CodigoCliente);
                }

                // Validar Item
                if (!string.IsNullOrEmpty(linha.CodigoItem) && !itensValidados.Contains(linha.CodigoItem))
                {
                    if (!ValidarItem(linha.CodigoItem))
                    {
                        validacao.AdicionarErro($"Item '{linha.CodigoItem}' não encontrado no SAP");
                    }
                    itensValidados.Add(linha.CodigoItem);
                }

                // Validar Utilização
                if (!string.IsNullOrEmpty(linha.Utilizacao) && !utilizacoesValidadas.Contains(linha.Utilizacao))
                {
                    if (!ValidarUtilizacao(linha.Utilizacao))
                    {
                        validacao.AdicionarAviso($"Utilização '{linha.Utilizacao}' não encontrada.");
                    }
                    utilizacoesValidadas.Add(linha.Utilizacao);
                }

                // Validar Código de Imposto
                if (!string.IsNullOrEmpty(linha.CodigoImposto) && !impostosValidados.Contains(linha.CodigoImposto))
                {
                    if (!ValidarCodigoImposto(linha.CodigoImposto))
                    {
                        validacao.AdicionarErro($"Código de imposto '{linha.CodigoImposto}' não encontrado no SAP");
                    }
                    impostosValidados.Add(linha.CodigoImposto);
                }

                // Validar Sequência NF
                if (!string.IsNullOrEmpty(linha.CodSeq) && !sequenciasValidadas.Contains(linha.CodSeq))
                {
                    if (!ValidarSequenciaNF(linha.CodSeq))
                    {
                        validacao.AdicionarAviso($"Sequência NF '{linha.CodSeq}' não encontrada.");
                    }
                    sequenciasValidadas.Add(linha.CodSeq);
                }

                // Validar Condição de Pagamento
                if (!string.IsNullOrEmpty(linha.CondicaoPagamento) && !condicoesValidadas.Contains(linha.CondicaoPagamento))
                {
                    if (!ValidarCondicaoPagamento(linha.CondicaoPagamento))
                    {
                        validacao.AdicionarAviso($"Condição de pagamento '{linha.CondicaoPagamento}' não encontrada.");
                    }
                    condicoesValidadas.Add(linha.CondicaoPagamento);
                }

                // Validar valor
                if (linha.Valor <= 0)
                {
                    validacao.AdicionarErro($"Linha {linha.NumeroLinha}: Valor deve ser maior que zero");
                }
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
            catch
            {
                return false;
            }
            finally
            {
                if (oBP != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
            }
        }

        private bool ValidarItem(string itemCode)
        {
            Items oItem = null;
            try
            {
                oItem = (Items)_company.GetBusinessObject(BoObjectTypes.oItems);
                return oItem.GetByKey(itemCode);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (oItem != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
            }
        }

        private bool ValidarUtilizacao(string usage)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = $@"SELECT ""ID"" FROM OUSG WHERE ""ID"" = '{usage}'";
                oRecordset.DoQuery(query);
                return !oRecordset.EoF;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        private bool ValidarCodigoImposto(string taxCode)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = $@"SELECT ""Code"" FROM OSTC WHERE ""Code"" = '{taxCode}'";
                oRecordset.DoQuery(query);
                return !oRecordset.EoF;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        private bool ValidarSequenciaNF(string seqCode)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = $@"SELECT ""SeqCode"" FROM NFN1 WHERE ""SeqCode"" = '{seqCode}' AND ""ObjectCode"" = '13'";
                oRecordset.DoQuery(query);
                return !oRecordset.EoF;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        private bool ValidarCondicaoPagamento(string groupNum)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = $@"SELECT ""GroupNum"" FROM OCTG WHERE ""GroupNum"" = '{groupNum}'";
                oRecordset.DoQuery(query);
                return !oRecordset.EoF;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        #endregion

        #region Métodos de Diagnóstico

        /// <summary>
        /// Corrige totais inconsistentes usando o método centralizado
        /// </summary>
        public bool CorrigirTotaisGrupo(string grupoCode)
        {
            try
            {
                return RecalcularStatusGrupoEficiente(grupoCode);
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }
}