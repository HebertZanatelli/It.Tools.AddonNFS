using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Text.Json;
using System.Configuration;
using System.Linq;
using System.Threading;

namespace ItTech.Tool.AddonNFS.Services
{
    /// <summary>
    /// ✅ ServiceLayerInvoiceClient ATUALIZADO - Usa ServiceLayerConnector que funciona
    /// Mantém toda a interface pública para compatibilidade
    /// </summary>
    public class ServiceLayerInvoiceClient : IDisposable
    {
        #region Properties and Fields

        private readonly ServiceLayerConnector _connector;
        private readonly bool _ownsConnector;
        private bool _disposed = false;

        private readonly CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();

        public bool IsConnected => ServiceLayerConnector.IsConnected;
        public string SessionId => ServiceLayerConnector.SessionId;

        public event EventHandler<InvoiceCreatedEventArgs> OnInvoiceCreated;
        public event EventHandler<InvoiceErrorEventArgs> OnInvoiceError;

        #endregion

        #region Constructors

        /// <summary>
        /// ✅ Construtor que cria seu próprio connector
        /// </summary>
        public ServiceLayerInvoiceClient(string serviceLayerUrl = null)
        {
            _connector = new ServiceLayerConnector(serviceLayerUrl);
            _ownsConnector = true;
        }

        /// <summary>
        /// ✅ Construtor que recebe connector existente
        /// </summary>
        public ServiceLayerInvoiceClient(ServiceLayerConnector connector)
        {
            _connector = connector ?? throw new ArgumentNullException(nameof(connector));
            _ownsConnector = false;
        }

        #endregion

        #region Connection Methods - USANDO SERVICELAYERCONNECTOR

        /// <summary>
        /// ✅ CONECTA usando ServiceLayerConnector (que funciona)
        /// </summary>
        public async Task<bool> ConnectAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(30)))
                using (var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken, timeoutCts.Token, _cancellationTokenSource.Token))
                {
                    return await _connector.ConnectAsync(combinedCts.Token).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
                OnInvoiceError?.Invoke(this, new InvoiceErrorEventArgs
                {
                    Operation = "Connect",
                    ErrorMessage = "Timeout ao conectar com Service Layer"
                });
                return false;
            }
            catch (Exception ex)
            {
                OnInvoiceError?.Invoke(this, new InvoiceErrorEventArgs
                {
                    Operation = "Connect",
                    Exception = ex,
                    ErrorMessage = $"Erro ao conectar: {ex.Message}"
                });
                return false;
            }
        }

        /// <summary>
        /// ✅ Conecta com parâmetros (compatibilidade)
        /// </summary>
        public async Task<bool> ConnectAsync(string companyDB, string username, string password,
            int? language = null, CancellationToken cancellationToken = default)
        {
            // Em addon, usar sessão do SAP B1 (ignora parâmetros)
            return await ConnectAsync(cancellationToken);
        }

        /// <summary>
        /// ✅ Desconecta
        /// </summary>
        public async Task<bool> DisconnectAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5)))
                using (var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken, timeoutCts.Token, _cancellationTokenSource.Token))
                {
                    return await _connector.LogoutAsync(combinedCts.Token).ConfigureAwait(false);
                }
            }
            catch
            {
                return false; // Timeout é considerado sucesso no logout
            }
        }

        /// <summary>
        /// ✅ Versões síncronas
        /// </summary>
        public bool Connect()
        {
            try
            {
                _connector.GetSession(); // Usar método síncrono direto
                return IsConnected;
            }
            catch
            {
                return false;
            }
        }

        public bool Disconnect()
        {
            try
            {
                return Task.Run(async () => await DisconnectAsync().ConfigureAwait(false))
                    .ConfigureAwait(false)
                    .GetAwaiter()
                    .GetResult();
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Invoice Operations - USANDO SERVICELAYERCONNECTOR.CALL

        /// <summary>
        /// ✅ Cria fatura usando ServiceLayerConnector.Call (método que funciona)
        /// </summary>
        public async Task<InvoiceResponse> CreateInvoiceAsync(InvoiceRequest invoice,
            CancellationToken cancellationToken = default)
        {
            try
            {
                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(60)))
                using (var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken, timeoutCts.Token, _cancellationTokenSource.Token))
                {
                    ValidateInvoiceRequest(invoice);

                    // ✅ GARANTIR que está conectado E renovar se necessário
                    if (!IsConnected || ServiceLayerConnector.NeedsRenewal())
                    {
                        Console.WriteLine($"[InvoiceClient] Conectando/renovando sessão...");
                        await ConnectAsync(combinedCts.Token).ConfigureAwait(false);

                        if (!IsConnected)
                        {
                            throw new ServiceLayerException("Falha ao conectar/renovar sessão Service Layer", -1, null);
                        }
                    }

                    // ✅ DOUBLE CHECK: Renovar explicitamente
                    _connector.RenewIfNeeded();

                    // ✅ PREPARAR JSON
                    var options = new JsonSerializerOptions
                    {

                        WriteIndented = false
                    };

                    var json = JsonSerializer.Serialize(invoice, options);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    // ✅ CRIAR REQUEST para ServiceLayerConnector.Call
                    var request = new HttpRequestMessage(HttpMethod.Post, "/b1s/v1/Invoices");
                    request.Content = content;

                    // ✅ USAR ServiceLayerConnector.Call (método que funciona)
                    var response = await ServiceLayerConnector.Call(request).ConfigureAwait(false);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync()
                            .ConfigureAwait(false);
                        var invoiceResponse = JsonSerializer.Deserialize<InvoiceResponse>(responseContent, options);

                        OnInvoiceCreated?.Invoke(this, new InvoiceCreatedEventArgs
                        {
                            Invoice = invoiceResponse,
                            RequestData = invoice,
                            StatusCode = response.StatusCode
                        });

                        return invoiceResponse;
                    }
                    else
                    {
                        var errorContent = await response.Content.ReadAsStringAsync()
                            .ConfigureAwait(false);
                        var error = TryParseServiceLayerError(errorContent);

                        OnInvoiceError?.Invoke(this, new InvoiceErrorEventArgs
                        {
                            Operation = "CreateInvoice",
                            StatusCode = response.StatusCode,
                            ErrorMessage = error.Message,
                            ErrorCode = error.Code,
                            RequestData = invoice,
                            ResponseContent = errorContent
                        });

                        throw new ServiceLayerException($"Erro ao criar fatura: {error.Message}",
                            error.Code, response.StatusCode);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw new ServiceLayerException("Timeout ao criar fatura", -1, null);
            }
            catch (ServiceLayerConnectionException ex)
            {
                throw new ServiceLayerException($"Erro de conexão: {ex.Message}", -1, null, ex);
            }
            catch (ServiceLayerException)
            {
                throw;
            }
            catch (Exception ex)
            {
                OnInvoiceError?.Invoke(this, new InvoiceErrorEventArgs
                {
                    Operation = "CreateInvoice",
                    Exception = ex,
                    ErrorMessage = ex.Message,
                    RequestData = invoice
                });

                throw new ServiceLayerException($"Erro inesperado ao criar fatura: {ex.Message}", -1, null, ex);
            }
        }

        /// <summary>
        /// ✅ Versão síncrona CORRIGIDA com renovação automática
        /// </summary>
        public InvoiceResponse CreateInvoice(InvoiceRequest invoice)
        {
            try
            {
                // ✅ GARANTIR conexão ANTES de processar
                if (!IsConnected || ServiceLayerConnector.NeedsRenewal())
                {
                    Console.WriteLine($"[InvoiceClient] Conectando/renovando sessão síncrono...");
                    bool conectou = Connect();
                    if (!conectou)
                    {
                        throw new ServiceLayerException("Falha ao conectar Service Layer", -1, null);
                    }
                }

                return Task.Run(async () => await CreateInvoiceAsync(invoice).ConfigureAwait(false))
                    .ConfigureAwait(false)
                    .GetAwaiter()
                    .GetResult();
            }
            catch (AggregateException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }

        /// <summary>
        /// ✅ Cria fatura simplificada
        /// </summary>
        public async Task<InvoiceResponse> CreateSimpleInvoiceAsync(
            string cardCode,
            DateTime docDate,
            DateTime docDueDate,
            string itemCode,
            decimal quantity,
            decimal unitPrice,
            string taxCode,
            int paymentGroupCode = -2,
            string openingRemarks = "BANCO XXXX",
            string tipTrib = "V",
            int usage = 13,
            int? sequenceCode = 31,
            string sequenceModel = "46",
            string bplId = "",
            CancellationToken cancellationToken = default)
        {
            var invoice = new InvoiceRequest
            {
                CardCode = cardCode,
                DocDate = docDate,
                PaymentGroupCode = paymentGroupCode,
                SequenceCode = sequenceCode,
                SequenceModel = sequenceModel,
                U_SKILL_TipTrib = tipTrib,
                OpeningRemarks = openingRemarks,
                BPL_IDAssignedToInvoice = bplId,
                DocumentLines = new List<InvoiceDocumentLine>
                {
                    new InvoiceDocumentLine
                    {
                        ItemCode = itemCode,
                        Quantity = quantity,
                        UnitPrice = unitPrice,
                        Usage = usage,
                        TaxCode = taxCode
                    }
                }
            };

            return await CreateInvoiceAsync(invoice, cancellationToken).ConfigureAwait(false);
        }

        #endregion

        #region Helper Methods - MANTIDOS

        private void ValidateInvoiceRequest(InvoiceRequest invoice)
        {
            if (invoice == null)
                throw new ArgumentNullException(nameof(invoice), "Dados da fatura são obrigatórios");

            var errors = new List<string>();

            if (string.IsNullOrWhiteSpace(invoice.CardCode))
                errors.Add("CardCode é obrigatório");

            if (string.IsNullOrWhiteSpace(invoice.U_SKILL_TipTrib))
                errors.Add("U_SKILL_TipTrib é obrigatório para emissão da nota");

            if (string.IsNullOrWhiteSpace(invoice.OpeningRemarks))
                errors.Add("OpeningRemarks é obrigatório");

            if (invoice.DocumentLines == null || !invoice.DocumentLines.Any())
                errors.Add("DocumentLines é obrigatório - pelo menos um item deve ser informado");

            if (invoice.DocumentLines != null)
            {
                for (int i = 0; i < invoice.DocumentLines.Count; i++)
                {
                    var line = invoice.DocumentLines[i];

                    if (string.IsNullOrWhiteSpace(line.ItemCode))
                        errors.Add($"DocumentLines[{i}].ItemCode é obrigatório");

                    if (line.Quantity <= 0)
                        errors.Add($"DocumentLines[{i}].Quantity deve ser maior que zero");

                    if (line.UnitPrice < 0)
                        errors.Add($"DocumentLines[{i}].UnitPrice não pode ser negativo");

                    if (!line.Usage.HasValue)
                        errors.Add($"DocumentLines[{i}].Usage é obrigatório");

                    if (string.IsNullOrWhiteSpace(line.TaxCode))
                        errors.Add($"DocumentLines[{i}].TaxCode é obrigatório");
                }
            }

            if (errors.Any())
                throw new ArgumentException($"Dados da fatura inválidos:\n- {string.Join("\n- ", errors)}");
        }

        private ServiceLayerError TryParseServiceLayerError(string errorContent)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var errorResponse = JsonSerializer.Deserialize<ServiceLayerErrorResponse>(errorContent, options);
                return errorResponse?.Error ?? new ServiceLayerError
                {
                    Message = "Erro desconhecido",
                    Code = -1
                };
            }
            catch
            {
                return new ServiceLayerError
                {
                    Message = errorContent ?? "Erro desconhecido",
                    Code = -1
                };
            }
        }

        #endregion

        #region Data Transfer Objects - MANTIDOS

        public class InvoiceRequest
        {
            public string CardCode { get; set; }
            public DateTime DocDate { get; set; } = DateTime.Now;
            public int PaymentGroupCode { get; set; } = -2;
            public int? SequenceCode { get; set; } = 31;
            public string SequenceModel { get; set; }
            public string U_SKILL_TipTrib { get; set; } = "V";
            public string OpeningRemarks { get; set; } = "BANCO XXXX";
            public string BPL_IDAssignedToInvoice { get; set; } = "";
            public InvoiceTaxExtension TaxExtension { get; set; } = null;
            public List<InvoiceDocumentLine> DocumentLines { get; set; } = new List<InvoiceDocumentLine>();

        }
        public class InvoiceTaxExtension
        {
            public string State { get; set; } = null;
            public string County { get; set; } = null;
        }

        public class InvoiceDocumentLine
        {
            public string ItemCode { get; set; }
            public decimal Quantity { get; set; } = 1;
            public decimal UnitPrice { get; set; }
            public int? Usage { get; set; } = 13;
            public string TaxCode { get; set; }
        }

        public class InvoiceResponse
        {
            public int DocEntry { get; set; }
            public int DocNum { get; set; }
            public string CardCode { get; set; }
            public DateTime DocDate { get; set; }
            public DateTime DocDueDate { get; set; }
            public decimal DocTotal { get; set; }
            public string OpeningRemarks { get; set; }
            public List<InvoiceDocumentLine> DocumentLines { get; set; }
        }

        private class ServiceLayerErrorResponse
        {
            public ServiceLayerError Error { get; set; }
        }

        private class ServiceLayerError
        {
            public string Message { get; set; }
            public int Code { get; set; }
        }

        #endregion

        #region Event Arguments - MANTIDOS

        public class InvoiceCreatedEventArgs : EventArgs
        {
            public InvoiceResponse Invoice { get; set; }
            public InvoiceRequest RequestData { get; set; }
            public HttpStatusCode StatusCode { get; set; }
        }

        public class InvoiceErrorEventArgs : EventArgs
        {
            public string Operation { get; set; }
            public HttpStatusCode? StatusCode { get; set; }
            public string ErrorMessage { get; set; }
            public int? ErrorCode { get; set; }
            public InvoiceRequest RequestData { get; set; }
            public string ResponseContent { get; set; }
            public Exception Exception { get; set; }
        }

        public class ServiceLayerException : Exception
        {
            public int ServiceLayerCode { get; }
            public HttpStatusCode? HttpStatusCode { get; }

            public ServiceLayerException(string message, int serviceLayerCode, HttpStatusCode? httpStatusCode, Exception innerException = null)
                : base(message, innerException)
            {
                ServiceLayerCode = serviceLayerCode;
                HttpStatusCode = httpStatusCode;
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                try
                {
                    _cancellationTokenSource?.Cancel();

                    if (_ownsConnector)
                    {
                        var logoutTask = DisconnectAsync();
                        if (!logoutTask.Wait(TimeSpan.FromSeconds(1)))
                        {
                            // Se não conseguir desconectar em 1 segundo, prosseguir
                        }

                        _connector?.Dispose();
                    }
                }
                catch
                {
                    // Ignorar todos os erros no dispose
                }
                finally
                {
                    _cancellationTokenSource?.Dispose();
                    _disposed = true;
                }
            }
        }

        #endregion
    }
}