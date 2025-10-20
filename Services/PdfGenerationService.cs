using System;
using System.Configuration;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace ItTech.Tool.AddonNFS.Services
{
    /// <summary>
    /// Serviço responsável pela geração de arquivos PDF a partir de layouts do SAP Business One.
    /// Utiliza o SAP Business One API Gateway para uma exportação robusta via web service.
    /// </summary>
    public class PdfGenerationService
    {
        private readonly Company _company; 

        
        public PdfGenerationService(Company company)
        {
            _company = company; 
        }

        /// <summary>
        /// Gera um PDF para um documento de Entrega de forma síncrona, chamando a versão assíncrona.
        /// </summary>
        /// <param name="docEntry">O DocEntry da Entrega.</param>
        /// <param name="pastaDestino">A pasta onde o PDF será salvo.</param>
        /// <param name="nomeArquivo">O nome do arquivo (sem extensão).</param>
        /// <returns>Uma tupla indicando sucesso e uma mensagem (caminho do arquivo ou erro).</returns>
        public (bool Success, string Message) GerarPdfDeEntrega(int docEntry, string pastaDestino, string nomeArquivo)
        {
            try
            {
                // Chama a versão assíncrona e aguarda o resultado.
                // .GetAwaiter().GetResult() é uma forma mais segura de aguardar Tasks em alguns contextos do que .Result.
                return GerarPdfViaApiGatewayAsync(docEntry, "15", pastaDestino, nomeArquivo).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                // Captura exceções que podem ser encapsuladas pela Task
                System.Diagnostics.Debug.WriteLine($"ERRO SÍNCRONO GerarPdfDeEntrega: {ex}");
                return (false, $"Exceção na chamada síncrona: {ex.Message}");
            }
        }


        /// <summary>
        /// Gera um PDF para um documento específico via API Gateway de forma assíncrona.
        /// </summary>
        /// <param name="docEntry">O DocEntry do documento.</param>
        /// <param name="objType">O Object Type do documento (ex: "15" para Entrega).</param>
        /// <param name="pastaDestino">A pasta onde o PDF será salvo.</param>
        /// <param name="nomeArquivo">O nome do arquivo (sem extensão).</param>
        /// <returns>Uma tupla indicando sucesso e uma mensagem (caminho do arquivo ou erro).</returns>
        public async Task<(bool Success, string Message)> GerarPdfViaApiGatewayAsync(int docEntry, string objType, string pastaDestino, string nomeArquivo)
        {
            // Validação de Parâmetros
            string layoutCode = ConfigurationManager.AppSettings["DeliveryNote_LayoutCode"];
            if (string.IsNullOrWhiteSpace(layoutCode))
            {
                return (false, "A chave 'DeliveryNote_LayoutCode' não está definida no App.config.");
            }

            if (string.IsNullOrWhiteSpace(pastaDestino))
            {
                return (false, "A pasta de destino do PDF não foi informada.");
            }

            try
            {
                if (!Directory.Exists(pastaDestino))
                {
                    Directory.CreateDirectory(pastaDestino);
                }
            }
            catch (Exception ex)
            {
                return (false, $"Não foi possível criar a pasta de destino '{pastaDestino}'. Erro: {ex.Message}");
            }

            string caminhoCompleto = Path.Combine(pastaDestino, $"{nomeArquivo}.pdf");

            try
            {
                // 1. Garantir que a sessão com a API Gateway está ativa
                await ApiGatewayConnector.EnsureApiGatewaySessionAsync();
                HttpClient apiClient = ApiGatewayConnector.Client;

                if (apiClient == null)
                {
                    return (false, "Não foi possível obter um cliente HTTP autenticado para a API Gateway.");
                }

                // 2. Montar o Payload para a API com Dockey@ e ObjectId@
                object[] payload = new object[] {
                    new {
                        name = "DocKey@", 
                        type = "xsd:string",
                        value = new[] { new[] { docEntry.ToString() } }
                    },
                    new {
                        name = "ObjectId@",
                        type = "xsd:decimal", // Tipo decimal conforme a sua especificação
                        value = new[] { new[] { decimal.Parse(objType) } } // Converte a string objType para decimal
                    }
                };
              

                string jsonPayload = JsonSerializer.Serialize(payload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                // 3. Montar a URL da Requisição (relativa à BaseAddress do HttpClient)
                string requestUrl = $"rs/vl/Export PDFData?DocCode={layoutCode}";

                // 4. Fazer a Chamada POST
                System.Diagnostics.Debug.WriteLine($"[PdfGenerationService] Enviando requisição para: {apiClient.BaseAddress}{requestUrl}");
                HttpResponseMessage response = await apiClient.PostAsync(requestUrl, content);

                // 5. Tratar a Resposta
                if (response.IsSuccessStatusCode)
                {
                    string base64Pdf = await response.Content.ReadAsStringAsync();
                    if (string.IsNullOrWhiteSpace(base64Pdf) || base64Pdf.StartsWith("{")) // Verifica se não é um JSON de erro
                    {
                        return (false, $"API Gateway retornou uma resposta inesperada ou vazia: {base64Pdf}");
                    }

                    // Decodificar Base64 para bytes e salvar no arquivo
                    byte[] pdfBytes = Convert.FromBase64String(base64Pdf);
                    File.WriteAllBytes(caminhoCompleto, pdfBytes);

                    System.Diagnostics.Debug.WriteLine($"[PdfGenerationService] PDF salvo com sucesso em: {caminhoCompleto}");
                    return (true, caminhoCompleto); // Sucesso, retorna o caminho completo do arquivo
                }
                else
                {
                    // Tratar resposta de erro
                    string errorContent = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"[PdfGenerationService] Erro da API Gateway ({response.StatusCode}): {errorContent}");
                    return (false, $"Erro da API Gateway ({response.StatusCode}): {errorContent}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfGenerationService] Exceção ao gerar PDF (DocEntry {docEntry}): {ex}");
                return (false, $"Exceção interna ao gerar PDF: {ex.Message}");
            }
        }
    }
}

