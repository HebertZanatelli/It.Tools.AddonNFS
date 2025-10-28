using System;
using System.Configuration;
using System.IO;
using SAPbobsCOM;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Diagnostics; // Para Debug.WriteLine

namespace ItTech.Tool.AddonNFS.Services
{
    /// <summary>
    /// Serviço responsável pela geração de arquivos PDF a partir de layouts do SAP Business One.
    /// Utiliza a exportação direta do Crystal Reports (local).
    /// </summary>
    public class PdfGenerationService
    {
        private readonly Company _company;

        // Construtor que aceita o objeto Company
        public PdfGenerationService(Company company)
        {
            _company = company ?? throw new ArgumentNullException(nameof(company));
        }

        /// <summary>
        /// Gera um PDF para um documento de Entrega (ODLN) usando Crystal Reports local.
        /// </summary>
        /// <param name="docEntry">O DocEntry da Entrega (ODLN).</param>
        /// <param name="pastaDestino">A pasta onde o PDF será salvo.</param>
        /// <param name="nomeArquivo">O nome do arquivo (sem extensão).</param>
        /// <returns>Uma tupla indicando sucesso e uma mensagem (caminho do arquivo ou erro).</returns>
        public (bool Success, string Message) GerarPdfDeEntrega(int docEntry, string pastaDestino, string nomeArquivo)
        {
            ReportDocument oRPT = new ReportDocument();
            string caminhoCompleto = "";
            string reportPath = "";
            // *** CORREÇÃO: Declarar a variável ANTES do try ***
            string nomeParametroDocEntry = "DocKey@"; // Nome padrão, confirme no seu .RPT

            // Log de início
            Debug.WriteLine($"[PdfGenerationService] Iniciando geração PDF para DocEntry: {docEntry}, Destino: {pastaDestino}, Arquivo: {nomeArquivo}");

            try
            {
                // --- 1. Validação e Preparação do Caminho ---
                if (string.IsNullOrWhiteSpace(pastaDestino))
                    return LogErrorAndReturn(docEntry, false, "Pasta de destino não informada.");

                if (string.IsNullOrWhiteSpace(nomeArquivo))
                    return LogErrorAndReturn(docEntry, false, "Nome do arquivo não informado.");

                if (!Directory.Exists(pastaDestino))
                {
                    try
                    {
                        Directory.CreateDirectory(pastaDestino);
                        Debug.WriteLine($"[PdfGenerationService] Pasta de destino criada: {pastaDestino}");
                    }
                    catch (Exception ex)
                    {
                        return LogErrorAndReturn(docEntry, false, $"Erro ao criar pasta de destino '{pastaDestino}': {ex.Message}");
                    }
                }
                caminhoCompleto = Path.Combine(pastaDestino, $"{nomeArquivo}.pdf");
                Debug.WriteLine($"[PdfGenerationService] Caminho completo do PDF: {caminhoCompleto}");

                // --- 2. Carregar o Arquivo .RPT ---
                string layoutCode = ConfigurationManager.AppSettings["DeliveryNote_LayoutCode"];
                if (string.IsNullOrWhiteSpace(layoutCode))
                    return LogErrorAndReturn(docEntry, false, "Chave 'DeliveryNote_LayoutCode' não encontrada ou vazia no App.config.");

                if (!layoutCode.EndsWith(".rpt", StringComparison.OrdinalIgnoreCase))
                    layoutCode += ".rpt";

                reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rpt", layoutCode);
                Debug.WriteLine($"[PdfGenerationService] Tentando carregar relatório: {reportPath}");

                if (!File.Exists(reportPath))
                    return LogErrorAndReturn(docEntry, false, $"Arquivo de relatório não encontrado: {reportPath}");

                oRPT.Load(reportPath);
                Debug.WriteLine($"[PdfGenerationService] Relatório '{layoutCode}' carregado.");

                // --- 3. Configurar Conexão com o Banco de Dados (Segura) ---
                string crystalUser = ConfigurationManager.AppSettings["CrystalReportDbUser"];
                string crystalPass = ConfigurationManager.AppSettings["CrystalReportDbPassword"]; // !! PROTEJA ESTA SENHA !!

                if (string.IsNullOrEmpty(crystalUser) || string.IsNullOrEmpty(crystalPass))
                    return LogErrorAndReturn(docEntry, false, "Credenciais 'CrystalReportDbUser' ou 'CrystalReportDbPassword' não configuradas/vazias no App.config.");

                Debug.WriteLine($"[PdfGenerationService] Configurando logon para DB: {_company.CompanyDB}, Server: {_company.Server}, User: {crystalUser}");
                oRPT.SetDatabaseLogon(crystalUser, crystalPass, _company.Server, _company.CompanyDB);
                foreach (ReportDocument subreport in oRPT.Subreports)
                {
                    subreport.SetDatabaseLogon(crystalUser, crystalPass, _company.Server, _company.CompanyDB);
                }

                // --- 4. Passar o Parâmetro DocEntry ---
                Debug.WriteLine($"[PdfGenerationService] Definindo parâmetro '{nomeParametroDocEntry}' = {docEntry}");
                try
                {
                    oRPT.SetParameterValue(nomeParametroDocEntry, docEntry);
                }
                catch (Exception exParam)
                {
                    // A variável nomeParametroDocEntry agora é acessível aqui
                    return LogErrorAndReturn(docEntry, false, $"Erro ao definir parâmetro '{nomeParametroDocEntry}'. Verifique se o parâmetro existe no relatório '{layoutCode}'. Erro Crystal: {exParam.Message}");
                }

                // --- 5. Exportar o PDF ---
                Debug.WriteLine($"[PdfGenerationService] Exportando para: {caminhoCompleto}");
                oRPT.ExportToDisk(ExportFormatType.PortableDocFormat, caminhoCompleto);

                Debug.WriteLine($"[PdfGenerationService] PDF salvo com sucesso: {caminhoCompleto}");
                return (true, caminhoCompleto); // Retorna sucesso e o caminho
            }
            catch (LogOnException logonEx) // Exceção específica para falha de logon
            {
                return LogErrorAndReturn(docEntry, false, $"Erro de logon no Crystal Reports. Verifique usuário/senha ('CrystalReportDbUser', 'CrystalReportDbPassword') no App.config e permissões no banco. Erro Crystal: {logonEx.Message}");
            }
            catch (ParameterFieldCurrentValueException paramEx) // Exceção específica para parâmetro
            {
                // *** CORREÇÃO: A variável agora está acessível aqui ***
                return LogErrorAndReturn(docEntry, false, $"Erro ao processar parâmetro '{nomeParametroDocEntry}' no Crystal. Verifique o tipo e valor esperado no relatório. Erro Crystal: {paramEx.Message}");
            }
            catch (Exception ex) // Captura geral
            {
                return LogErrorAndReturn(docEntry, false, $"Erro inesperado ao gerar PDF Crystal: {ex.Message} (Relatório: {reportPath})");
            }
            finally
            {
                // --- 6. Liberar Recursos (MUITO IMPORTANTE) ---
                if (oRPT != null)
                {
                    try { oRPT.Close(); } catch { }
                    try { oRPT.Dispose(); } catch { }
                    Debug.WriteLine($"[PdfGenerationService] Recursos do Crystal Reports liberados para DocEntry: {docEntry}");
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Método auxiliar para log e retorno padronizado de erro
        private (bool Success, string Message) LogErrorAndReturn(int docEntry, bool success, string message)
        {
            Debug.WriteLine($"[PdfGenerationService] {(success ? "INFO" : "ERRO")} (DocEntry: {docEntry}): {message}");
            return (success, message);
        }
    }
}