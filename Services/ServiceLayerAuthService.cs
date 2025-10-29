using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;
using SAPbouiCOM.Framework;

namespace ItTech.Tool.AddonNFS.Services
{
    /// <summary>
    /// ✅ ServiceLayerConnector - BASEADO NO CÓDIGO QUE FUNCIONA NA EMPRESA
    /// Usa GetServiceLayerConnectionContext exatamente como no exemplo
    /// </summary>
    public class ServiceLayerConnector : IDisposable
    {
        #region Fields

        private static HttpClient _client = null;
        private static string _currentSessionId = null;
        private static string _currentDatabase = null;
        private static Uri _serviceLayerUrl = null;
        private static DateTime _sessionExpiry = DateTime.MinValue;

        private readonly string _baseUrl;
        private bool _disposed = false;

        // Lock para thread-safety
        private static readonly object _lock = new object();

        #endregion

        #region Constructor

        public ServiceLayerConnector(string serviceLayerUrl = null)
        {
            // ✅ URL SEM BARRA NO FINAL (como no exemplo que funciona)
            _baseUrl = serviceLayerUrl ??
                      ConfigurationManager.AppSettings["ServiceLayerUrl"] ??
                      "https://hanab1:50000/b1s/v1";

            // Remover barra no final se existir
            if (_baseUrl.EndsWith("/"))
            {
                _baseUrl = _baseUrl.TrimEnd('/');
            }
        }

        #endregion

        #region Public Properties

        public static bool IsConnected => _client != null && _currentSessionId != null && DateTime.Now < _sessionExpiry;
        public static string SessionId => _currentSessionId;
        public static string Database => _currentDatabase;
        public static HttpClient Client => _client;

        #endregion

        #region Connection Methods - EXATAMENTE COMO O EXEMPLO QUE FUNCIONA

        /// <summary>
        /// ✅ GetSession - IMPLEMENTAÇÃO EXATA DO CÓDIGO QUE FUNCIONA
        /// </summary>
        public void GetSession()
        {
            lock (_lock)
            {
                try
                {
                    string session = string.Empty;
                    string nomeBanco = string.Empty;

                    //  EXATAMENTE como no exemplo - SEM barra no final
                    string serviceLayerAddress = _baseUrl;

                    //  OBTER contexto do SAP B1 - MÉTODO QUE FUNCIONA
                    string context = Application.SBO_Application.Company.GetServiceLayerConnectionContext(serviceLayerAddress);

                    if (string.IsNullOrEmpty(context))
                    {
                        throw new ServiceLayerConnectionException("Falha ao obter contexto do Service Layer. Verifique se está logado no SAP B1.");
                    }

                    //  URL como no exemplo
                    Uri UrlServiceLayer = new Uri(serviceLayerAddress + "/");

                    //  EXTRAIR sessão do contexto
                    string[] cookieItems = context.Split(';');
                    foreach (var cookieItem in cookieItems)
                    {
                        string[] parts = cookieItem.Split('=');
                        if (parts.Length > 0)
                        {
                            if (parts[0].Trim() == "B1SESSION")
                            {
                                session = parts[1].Trim();
                                break; 
                            }
                        }
                    }

                    if (string.IsNullOrEmpty(session))
                    {
                        throw new ServiceLayerConnectionException("B1SESSION não encontrado no contexto");
                    }

                    
                    nomeBanco = Application.SBO_Application.Company.DatabaseName;

                    if (string.IsNullOrEmpty(nomeBanco))
                    {
                        throw new ServiceLayerConnectionException("Database não encontrado na Application");
                    }

                    
                    SetClient(UrlServiceLayer, nomeBanco, session);

                    
                    _sessionExpiry = DateTime.Now.AddMinutes(15);

                    System.Diagnostics.Debug.WriteLine($"[ServiceLayer] ✅ Sessão obtida com sucesso!");
                    System.Diagnostics.Debug.WriteLine($"[ServiceLayer] SessionId: {session}");
                    System.Diagnostics.Debug.WriteLine($"[ServiceLayer] Database: {nomeBanco}");
                    System.Diagnostics.Debug.WriteLine($"[ServiceLayer] URL: {serviceLayerAddress}");
                }
                catch (Exception ex)
                {
                    throw new ServiceLayerConnectionException($"Erro ao obter sessão: {ex.Message}", ex);
                }
            }
        }

        /// <summary>
        ///  SetClient - IMPLEMENTAÇÃO  com debug detalhado
        /// </summary>
        private static void SetClient(Uri UrlServiceLayer, string database, string sessionId)
        {
            try
            {
                // ✅ Dispose do cliente anterior se existir
                _client?.Dispose();

                Console.WriteLine($"[ServiceLayer] Configurando HttpClient...");
                Console.WriteLine($"[ServiceLayer] URL: {UrlServiceLayer}");
                Console.WriteLine($"[ServiceLayer] Database: {database}");
                Console.WriteLine($"[ServiceLayer] SessionId: {sessionId}");

                // ✅ CONFIGURAÇÃO EXATA como no exemplo que funciona
                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = new CookieContainer();

                // ✅ ADICIONAR cookies na ordem correta
                var companyDbCookie = new Cookie("CompanyDB", database, "/b1s/v1");
                var sessionCookie = new Cookie("B1SESSION", sessionId, "/b1s/v1");
                var routeCookie = new Cookie("ROUTEID", ".node0", "/b1s");

                handler.CookieContainer.Add(UrlServiceLayer, companyDbCookie);
                handler.CookieContainer.Add(UrlServiceLayer, sessionCookie);
                handler.CookieContainer.Add(UrlServiceLayer, routeCookie);

                Console.WriteLine($"[ServiceLayer] Cookies adicionados:");
                Console.WriteLine($"[ServiceLayer] - CompanyDB: {database}");
                Console.WriteLine($"[ServiceLayer] - B1SESSION: {sessionId}");
                Console.WriteLine($"[ServiceLayer] - ROUTEID: .node0");

                // ✅ SSL bypass
                handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => { return true; };

                // ✅ CRIAR cliente
                _client = new HttpClient(handler);
                _client.BaseAddress = UrlServiceLayer;
                _client.DefaultRequestHeaders.ExpectContinue = false;
                _client.DefaultRequestHeaders.Add("Cache-Control", "no-cache, no-store, max-age=0, must-revalidate");
                _client.DefaultRequestHeaders.Add("SessionId", sessionId);

                // ✅ Timeout mais longo para operações
                _client.Timeout = TimeSpan.FromMinutes(2);

                // Armazenar dados da sessão
                _currentSessionId = sessionId;
                _currentDatabase = database;
                _serviceLayerUrl = UrlServiceLayer;

                Console.WriteLine($"[ServiceLayer] ✅ HttpClient configurado com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ServiceLayer] ❌ Erro ao configurar HttpClient: {ex.Message}");
                throw new ServiceLayerConnectionException($"Erro ao configurar HttpClient: {ex.Message}", ex);
            }
        }

        #endregion

        #region Call Method - COMO NO EXEMPLO

        /// <summary>
        /// ✅ Call - IMPLEMENTAÇÃO CORRIGIDA com renovação automática de sessão
        /// </summary>
        public static async Task<HttpResponseMessage> Call(HttpRequestMessage request)
        {
            Console.WriteLine($"[ServiceLayer] Call iniciado");

            // ✅ VERIFICAR e RENOVAR sessão se necessário
            if (_client == null || NeedsRenewal() || !IsConnected)
            {
                Console.WriteLine($"[ServiceLayer] Renovando sessão...");

                // Criar instância temporária para renovar
                var tempConnector = new ServiceLayerConnector();
                tempConnector.GetSession();

                if (_client == null)
                {
                    throw new ServiceLayerConnectionException("Falha ao renovar sessão do Service Layer!");
                }
            }

            try
            {
                // ✅ DEBUG: Verificar cookies antes da requisição
                Console.WriteLine($"[ServiceLayer] SessionId: {_currentSessionId}");
                Console.WriteLine($"[ServiceLayer] Database: {_currentDatabase}");
                Console.WriteLine($"[ServiceLayer] URL: {request.RequestUri}");

                HttpResponseMessage response = await _client.SendAsync(request);

                // ✅ VERIFICAR se a resposta indica sessão inválida
                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"[ServiceLayer] Erro: {errorContent}");

                    if (errorContent.Contains("Invalid session") || errorContent.Contains("session already timeout"))
                    {
                        Console.WriteLine($"[ServiceLayer] Sessão inválida detectada - limpando...");
                        ClearSession();
                        throw new ServiceLayerConnectionException("Sessão inválida - será renovada na próxima requisição");
                    }
                }

                return response;
            }
            catch (Exception er)
            {
                Console.WriteLine($"[ServiceLayer] Erro na requisição: {er.Message}");

                if (er.Message.IndexOf("Invalid session") != -1 || er.Message.IndexOf("session already timeout") != -1)
                {
                    Console.WriteLine($"[ServiceLayer] Limpando sessão inválida...");
                    ClearSession();
                }
                throw er;
            }
        }

        #endregion

        #region Async Connection Methods - Para compatibilidade

        /// <summary>
        /// ✅ Versão async para compatibilidade com ServiceLayerInvoiceClient
        /// </summary>
        public async Task<bool> ConnectAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                await Task.Run(() => GetSession(), cancellationToken);
                return IsConnected;
            }
            catch (OperationCanceledException)
            {
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// ✅ Login async para compatibilidade
        /// </summary>
        public async Task<bool> LoginAsync(string companyDB, string userName, string password, int? language = null, CancellationToken cancellationToken = default)
        {
            // Em addon, ignorar parâmetros e usar sessão do SAP B1
            return await ConnectAsync(cancellationToken);
        }

        /// <summary>
        /// ✅ Logout async para compatibilidade
        /// </summary>
        public async Task<bool> LogoutAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                await Task.Run(() => ClearSession(), cancellationToken);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// ✅ HttpClient autenticado para compatibilidade
        /// </summary>
        public async Task<HttpClient> GetAuthenticatedClientAsync(string companyDB = null, string userName = null, string password = null, CancellationToken cancellationToken = default)
        {
            if (!IsConnected)
            {
                await ConnectAsync(cancellationToken);
            }

            if (_client == null)
            {
                throw new ServiceLayerConnectionException("HttpClient não foi criado. Execute ConnectAsync() primeiro.");
            }

            return _client;
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// ✅ Limpar sessão COM DEBUG
        /// </summary>
        private static void ClearSession()
        {
            lock (_lock)
            {
                Console.WriteLine($"[ServiceLayer] Limpando sessão...");

                try
                {
                    _client?.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ServiceLayer] Erro ao dispose HttpClient: {ex.Message}");
                }

                _client = null;
                _currentSessionId = null;
                _currentDatabase = null;
                _sessionExpiry = DateTime.MinValue;

                Console.WriteLine($"[ServiceLayer] ✅ Sessão limpa");
            }
        }

        /// <summary>
        ///  Verificar se precisa renovar sessão (mais conservador)
        /// </summary>
        public static bool NeedsRenewal()
        {
            //  Renovar se a sessão expira em menos de 5 minutos OU já expirou
            return _sessionExpiry == DateTime.MinValue ||
                   DateTime.Now.AddMinutes(5) >= _sessionExpiry ||
                   string.IsNullOrEmpty(_currentSessionId);
        }

        /// <summary>
        ///  Renovar sessão se necessário (com debug)
        /// </summary>
        public void RenewIfNeeded()
        {
            if (!IsConnected || NeedsRenewal())
            {
                Console.WriteLine($"[ServiceLayer] Renovando sessão - Atual expira em: {_sessionExpiry}");
                GetSession();
            }
        }

        /// <summary>
        ///  Testar conectividade COM DEBUG DETALHADO
        /// </summary>
        public async Task<bool> TestConnectionAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                Console.WriteLine($"[ServiceLayer] Testando conectividade...");

                if (!IsConnected || NeedsRenewal())
                {
                    Console.WriteLine($"[ServiceLayer] Renovando sessão para teste...");
                    await ConnectAsync(cancellationToken);
                }

                if (_client == null)
                {
                    Console.WriteLine($"[ServiceLayer] ❌ HttpClient é null");
                    return false;
                }

                Console.WriteLine($"[ServiceLayer] Fazendo requisição de teste...");
                var request = new HttpRequestMessage(HttpMethod.Get, "/b1s/v1/$metadata");
                var response = await Call(request);

                bool sucesso = response.IsSuccessStatusCode;
                Console.WriteLine($"[ServiceLayer] Teste resultado: {(sucesso ? "✅ SUCESSO" : "❌ FALHA")} - Status: {response.StatusCode}");

                if (!sucesso)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"[ServiceLayer] Erro: {content}");
                }

                return sucesso;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ServiceLayer] ❌ Erro no teste: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Properties para compatibilidade

        public bool IsAuthenticated => IsConnected;
        public string SessionId_Property => SessionId;
        public string RouteId => ".node0"; // Fixo como no exemplo
        public DateTime ExpiresAt => _sessionExpiry;

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
                    ClearSession();
                }
                catch { }

                _disposed = true;
            }
        }

        #endregion
    }

    #region Custom Exception

    /// <summary>
    /// Exceção específica para erros de conexão do Service Layer
    /// </summary>
    public class ServiceLayerConnectionException : Exception
    {
        public ServiceLayerConnectionException(string message) : base(message) { }
        public ServiceLayerConnectionException(string message, Exception innerException) : base(message, innerException) { }
    }

    #endregion
}