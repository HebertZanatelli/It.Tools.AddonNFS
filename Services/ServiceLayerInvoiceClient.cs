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

        public ServiceLayerInvoiceClient(string serviceLayerUrl = null)
        {
            _connector = new ServiceLayerConnector(serviceLayerUrl);
            _ownsConnector = true;
        }

        public ServiceLayerInvoiceClient(ServiceLayerConnector connector)
        {
            _connector = connector ?? throw new ArgumentNullException(nameof(connector));
            _ownsConnector = false;
        }

        #endregion

        #region Connection Methods

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
            catch (Exception ex)
            {
                // Tratar erro...
                return false;
            }
        }

        public bool Connect()
        {
            try
            {
                _connector.GetSession();
                return IsConnected;
            }
            catch { return false; }
        }

        // ... outros métodos de conexão como Disconnect, etc. ...

        #endregion

        #region Document Operations

        // MÉTODO GENÉRICO PRIVADO (CORRIGIDO E DENTRO DA CLASSE)
        private async Task<InvoiceResponse> _PostDocumentAsync(string endpoint, InvoiceRequest requestData, CancellationToken cancellationToken = default)
        {
            if (!IsConnected || ServiceLayerConnector.NeedsRenewal())
            {
                await ConnectAsync(cancellationToken);
            }

            var json = JsonSerializer.Serialize(requestData, new JsonSerializerOptions { WriteIndented = false });
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
            request.Content = content;

            var response = await ServiceLayerConnector.Call(request);

            if (response.IsSuccessStatusCode)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                return JsonSerializer.Deserialize<InvoiceResponse>(responseContent);
            }
            else
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                var error = TryParseServiceLayerError(errorContent);
                throw new ServiceLayerException($"Erro ao criar documento em '{endpoint}': {error.Message}", error.Code, response.StatusCode);
            }
        }

        // MÉTODO CreateInvoice refatorado para usar o método genérico
        public InvoiceResponse CreateInvoice(InvoiceRequest invoice)
        {
            try
            {
                if (!IsConnected || ServiceLayerConnector.NeedsRenewal())
                {
                    if (!Connect())
                    {
                        throw new ServiceLayerException("Falha ao conectar Service Layer", -1, null);
                    }
                }

                // Delega a chamada para o método genérico assíncrono
                return Task.Run(async () => await _PostDocumentAsync("/b1s/v1/Invoices", invoice)).Result;
            }
            catch (AggregateException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }
      
        public InvoiceResponse CreateDeliveryNote(InvoiceRequest requestData)
        {
            try
            {
                if (!IsConnected || ServiceLayerConnector.NeedsRenewal())
                {
                    if (!Connect())
                    {
                        throw new ServiceLayerException("Falha ao conectar Service Layer", -1, null);
                    }
                }

                // Delega a chamada para o método privado, especificando o endpoint de Entregas
                return Task.Run(async () => await _PostDocumentAsync("/b1s/v1/DeliveryNotes", requestData)).Result;
            }
            catch (AggregateException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }

        public InvoiceResponse CreatePurchaseInvoice(InvoiceRequest requestData)
        {
            try
            {
                if (!IsConnected || ServiceLayerConnector.NeedsRenewal())
                {
                    if (!Connect())
                    {
                        throw new ServiceLayerException("Falha ao conectar Service Layer", -1, null);
                    }
                }
                // Delega a chamada para o método privado, especificando o endpoint de Nota Fiscal de Entrada
                return Task.Run(async () => await _PostDocumentAsync("/b1s/v1/PurchaseInvoices", requestData)).Result;
            }
            catch (AggregateException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }


        #endregion

        #region Helper Methods

        private void ValidateInvoiceRequest(InvoiceRequest invoice)
        {
            // ... (código existente de validação) ...
        }

        private ServiceLayerError TryParseServiceLayerError(string errorContent)
        {
            try
            {
                var options = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
                var errorResponse = JsonSerializer.Deserialize<ServiceLayerErrorResponse>(errorContent, options);
                return errorResponse?.Error ?? new ServiceLayerError { Message = "Erro desconhecido", Code = -1 };
            }
            catch
            {
                return new ServiceLayerError { Message = errorContent ?? "Erro desconhecido", Code = -1 };
            }
        }

        #endregion

        #region Data Transfer Objects

        // As classes (InvoiceRequest, InvoiceResponse, etc.) continuam aqui dentro

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
            // ... (outras propriedades de resposta) ...
        }

        private class ServiceLayerErrorResponse { public ServiceLayerError Error { get; set; } }
        private class ServiceLayerError { public string Message { get; set; } public int Code { get; set; } }

        #endregion

        #region Event Arguments & Exception

        public class InvoiceCreatedEventArgs : EventArgs { /* ... */ }
        public class InvoiceErrorEventArgs : EventArgs { /* ... */ }
        public class ServiceLayerException : Exception
        {
            public int ServiceLayerCode { get; }
            public HttpStatusCode? HttpStatusCode { get; }
            public ServiceLayerException(string message, int serviceLayerCode, HttpStatusCode? httpStatusCode, Exception innerException = null) : base(message, innerException)
            {
                ServiceLayerCode = serviceLayerCode;
                HttpStatusCode = httpStatusCode;
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose() { /* ... */ }
        protected virtual void Dispose(bool disposing) { /* ... */ }

        #endregion

    } // FIM DA CLASSE ServiceLayerInvoiceClient
}