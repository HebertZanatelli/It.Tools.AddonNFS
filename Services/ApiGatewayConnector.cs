//using System;
//using System.Configuration;
//using SAPbouiCOM;
//using System.Net;
//using System.Net.Http;
//using System.Text;
//using System.Text.Json;
//using System.Threading;
//using System.Threading.Tasks;
//using SAPbouiCOM.Framework; // Para Application.SBO_Application

//namespace ItTech.Tool.AddonNFS.Services
//{
//    /// <summary>
//    /// Gerencia a conexão e autenticação com o SAP Business One API Gateway.
//    /// Utiliza o endpoint /login e gerencia a sessão via Cookies.
//    /// Classe estática para manter uma única instância do HttpClient e estado de sessão.
//    /// </summary>
//    public static class ApiGatewayConnector // Usando static para simplicidade singleton
//    {
//        #region Campos Estáticos Privados

//        private static HttpClient _apiGatewayClient = null;
//        // CookieContainer é crucial para que o HttpClient gerencie automaticamente os cookies de sessão
//        private static readonly CookieContainer _cookieContainer = new CookieContainer();
//        private static Uri _apiGatewayBaseUri = null;
//        private static DateTime _apiGatewayExpiry = DateTime.MinValue;
//        private static readonly object _lockApiGateway = new object(); // Lock para thread-safety
//        private static bool _isDisposed = false; // Controle para Dispose
//        private static SemaphoreSlim _loginSemaphore = new SemaphoreSlim(1, 1); // Garante que apenas um login ocorra por vez

//        #endregion

//        #region Propriedades Públicas Estáticas

//        /// <summary>
//        /// Obtém a instância do HttpClient configurada e autenticada (se a sessão estiver ativa).
//        /// Retorna null se não estiver conectado. Chame EnsureApiGatewaySessionAsync antes de usar.
//        /// </summary>
//        public static HttpClient Client
//        {
//            get
//            {
//                lock (_lockApiGateway)
//                {
//                    // Retorna o cliente apenas se a sessão for considerada válida e não descartado
//                    return IsConnected ? _apiGatewayClient : null;
//                }
//            }
//        }

//        /// <summary>
//        /// Verifica se a conexão com a API Gateway está ativa e a sessão não expirou.
//        /// </summary>
//        public static bool IsConnected
//        {
//            get
//            {
//                lock (_lockApiGateway)
//                {
//                    // Verifica se não foi descartado, se o cliente existe e se a sessão é válida
//                    return !_isDisposed && _apiGatewayClient != null && DateTime.Now < _apiGatewayExpiry;
//                }
//            }
//        }

//        #endregion

//        #region Métodos Públicos Estáticos

//        /// <summary>
//        /// Garante que existe uma sessão ativa com a API Gateway.
//        /// Se a sessão não existir ou tiver expirado, tenta realizar o login de forma segura (evitando logins simultâneos).
//        /// Lança uma exceção se o login falhar.
//        /// </summary>
//        /// <param name="cancellationToken">Token de cancelamento.</param>
//        /// <exception cref="Exception">Lançada se o login falhar.</exception>
//        /// <exception cref="ObjectDisposedException">Lançada se o conector já foi descartado.</exception>
//        public static async Task EnsureApiGatewaySessionAsync(CancellationToken cancellationToken = default)
//        {
//            // Verificação rápida fora do lock (otimização)
//            if (IsConnected) return;

//            // Entra na seção crítica para verificar novamente e tentar o login se necessário
//            await _loginSemaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
//            try
//            {
//                // Verifica novamente dentro do lock do semáforo, caso outra thread já tenha conectado
//                if (IsConnected) return;

//                lock (_lockApiGateway)
//                {
//                    if (_isDisposed) throw new ObjectDisposedException(nameof(ApiGatewayConnector));
//                }

//                System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão API Gateway inválida. Tentando login...");
//                bool loginSuccess = await LoginApiGatewayAsyncInternal(cancellationToken).ConfigureAwait(false); // Chama método interno

//                if (!loginSuccess)
//                {
//                    // LoginApiGatewayAsyncInternal já logou o erro específico.
//                    throw new Exception("Falha ao estabelecer ou renovar a sessão com a API Gateway.");
//                }
//                System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão estabelecida/renovada.");
//            }
//            finally
//            {
//                _loginSemaphore.Release(); // Libera o semáforo
//            }
//        }


//        /// <summary>
//        /// Realiza o logout da API Gateway (best-effort) e limpa a sessão local.
//        /// </summary>
//        public static async Task LogoutAsync(CancellationToken cancellationToken = default)
//        {
//            HttpClient clientToLogout = null;
//            lock (_lockApiGateway)
//            {
//                clientToLogout = _apiGatewayClient; // Pega a referência atual para usar fora do lock
//            }

//            if (clientToLogout != null)
//            {
//                try
//                {
//                    // Tenta chamar o endpoint /logout
//                    // Ignoramos erros aqui, o principal é limpar localmente
//                    await clientToLogout.PostAsync("logout", null, cancellationToken).ConfigureAwait(false);
//                    System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Chamada /logout enviada (resultado ignorado).");
//                }
//                catch (Exception ex)
//                {
//                    System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Exceção durante /logout (ignorado): {ex.Message}");
//                }
//            }

//            ClearSession(); // Limpa o estado local
//            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão local e HttpClient limpos.");
//        }


//        /// <summary>
//        /// Libera os recursos (HttpClient). Deve ser chamado ao descarregar o Add-on.
//        /// </summary>
//        public static void Dispose()
//        {
//            lock (_lockApiGateway) // Garante exclusividade ao descartar
//            {
//                if (_isDisposed) return; // Já descartado

//                ClearSessionInternal(); // Limpa o cliente HTTP e estado
//                _isDisposed = true; // Marca como descartado
//            }
//            _loginSemaphore.Dispose(); // Libera o semáforo
//            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Recursos liberados.");
//        }

//        #endregion

//        #region Métodos Privados Estáticos

//        /// <summary>
//        /// Lógica interna de login. NÃO chamar diretamente, use EnsureApiGatewaySessionAsync.
//        /// Assume que já está dentro de um lock ou semáforo apropriado.
//        /// Retorna true se o login for bem-sucedido, false caso contrário.
//        /// </summary>
//        private static async Task<bool> LoginApiGatewayAsyncInternal(CancellationToken cancellationToken = default)
//        {
//            // Limpa estado anterior antes de tentar novo login (dentro do lock/semáforo)
//            ClearSessionInternal();

//            string apiUrl = null;
//            string user = null;
//            string pass = null;
//            string companyDB = null;

//            try
//            {
//                apiUrl = ConfigurationManager.AppSettings["ApiGatewayUrl"];
//                user = ConfigurationManager.AppSettings["SAP_Username"];
//                pass = ConfigurationManager.AppSettings["SAP_Password"];
//                companyDB = SAPbouiCOM.Framework.Application.SBO_Application.Company.DatabaseName; // Banco da sessão atual

//                if (string.IsNullOrEmpty(apiUrl) || string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass) || string.IsNullOrEmpty(companyDB))
//                {
//                    LogAndStatusBarError("Configurações ApiGatewayUrl/User/Password/CompanyDB ausentes ou inválidas.");
//                    return false;
//                }

//                // Prepara URL base
//                if (apiUrl.EndsWith("/")) apiUrl = apiUrl.TrimEnd('/');
//                _apiGatewayBaseUri = new Uri(apiUrl + "/"); // BaseAddress do HttpClient

//                // Prepara payload
//                var loginPayload = new { CompanyDB = companyDB, UserName = user, Password = pass };
//                string jsonPayload = JsonSerializer.Serialize(loginPayload);
//                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

//                // Configura HttpClientHandler
//                HttpClientHandler handler = new HttpClientHandler
//                {
//                    CookieContainer = _cookieContainer,
//                    ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true // Ignora SSL inválido
//                };

//                // Cria o cliente HTTP
//                var client = new HttpClient(handler) { BaseAddress = _apiGatewayBaseUri };
//                client.Timeout = TimeSpan.FromMinutes(2);

//                // Log da requisição
//                System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] POST -> {client.BaseAddress}login | Payload: {jsonPayload}");

//                try
//                {
//                    // Mede tempo da requisição
//                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
//                    HttpResponseMessage response = await client.PostAsync("login", content, cancellationToken).ConfigureAwait(false);
//                    stopwatch.Stop();

//                    string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

//                    System.Diagnostics.Debug.WriteLine(
//                        $"[ApiGatewayConnector] Response ({(int)response.StatusCode}) {response.StatusCode} | " +
//                        $"Tempo: {stopwatch.ElapsedMilliseconds}ms | Body: {responseBody}"
//                    );

//                    if (response.IsSuccessStatusCode)
//                    {
//                        int timeoutMinutes = 30; // Default
//                        try
//                        {
//                            using (JsonDocument document = JsonDocument.Parse(responseBody))
//                            {
//                                if (document.RootElement.TryGetProperty("SessionTimeout", out JsonElement timeoutElement) &&
//                                    timeoutElement.TryGetInt32(out int parsedTimeout))
//                                {
//                                    timeoutMinutes = parsedTimeout;
//                                }
//                            }
//                        }
//                        catch
//                        {
//                            // Ignora erro no parse, usa timeout padrão
//                        }

//                        // Define expiração com margem de segurança
//                        _apiGatewayExpiry = DateTime.Now.AddMinutes(Math.Max(1, timeoutMinutes - 1));

//                        // Guarda o cliente válido
//                        lock (_lockApiGateway)
//                        {
//                            _apiGatewayClient = client;
//                        }

//                        System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Login bem-sucedido. Sessão expira em: {_apiGatewayExpiry}");
//                        return true;
//                    }
//                    else
//                    {
//                        LogAndStatusBarError($"Falha no login da API Gateway ({response.StatusCode}): {responseBody}");
//                        client.Dispose();
//                        return false;
//                    }
//                }
//                catch (TaskCanceledException tex)
//                {
//                    LogAndStatusBarError($"Timeout ou cancelamento no login da API Gateway ({_apiGatewayBaseUri}login): {tex.Message}");
//                    client.Dispose();
//                    return false;
//                }
//                catch (Exception ex)
//                {
//                    LogAndStatusBarError($"Erro inesperado durante a chamada de login: {ex.Message}");
//                    client.Dispose();
//                    return false;
//                }
//            }
//            catch (Exception ex)
//            {
//                LogAndStatusBarError($"Exceção durante login na API Gateway: {ex.Message}");
//                lock (_lockApiGateway)
//                {
//                    _apiGatewayClient?.Dispose();
//                    _apiGatewayClient = null;
//                }
//                _apiGatewayExpiry = DateTime.MinValue;
//                return false;
//            }
//        }


//        /// <summary>
//        /// Limpa o estado da sessão e descarta o HttpClient. Deve ser chamado dentro de um lock.
//        /// </summary>
//        private static void ClearSessionInternal()
//        {
//            _apiGatewayClient?.Dispose(); // Descarta o cliente HTTP se existir
//            _apiGatewayClient = null;
//            _apiGatewayExpiry = DateTime.MinValue;
//            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Estado interno da sessão limpo.");
//        }

//        /// <summary>
//        /// Método auxiliar para limpar a sessão de forma segura (usado por Logout e Dispose).
//        /// </summary>
//        private static void ClearSession()
//        {
//            lock (_lockApiGateway)
//            {
//                ClearSessionInternal();
//            }
//        }


//        /// <summary>
//        /// Helper para logar erro no Debug e mostrar na StatusBar do B1.
//        /// </summary>
//        private static void LogAndStatusBarError(string message)
//        {
//            System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] ERRO: {message}");
//            try
//            {
//                SAPbouiCOM.Framework.Application.SBO_Application?.StatusBar?.SetText($"Erro API Gateway: {message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            catch { }
//        }

//        #endregion
//    }
//}
using System;
using System.Configuration;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace ItTech.Tool.AddonNFS.Services
{
    public static class ApiGatewayConnector
    {
        #region Campos privados

        private static readonly object _lock = new object();
        private static HttpClient _client = null;
        private static Uri _baseUri = null;
        private static DateTime _sessionExpiry = DateTime.MinValue;
        private static readonly CookieContainer _cookies = new CookieContainer();

        private static bool _isLoggingIn = false;
        private static bool _isDisposed = false;

        #endregion

        #region Propriedades públicas

        public static bool IsConnected => _client != null && DateTime.Now < _sessionExpiry && !_isDisposed;
        public static HttpClient Client => _client;
        public static DateTime Expiry => _sessionExpiry;

        #endregion

        #region Métodos principais

        /// <summary>
        /// Garante que exista uma sessão válida na API Gateway.
        /// </summary>
        public static async Task EnsureSessionAsync(CancellationToken cancellationToken = default)
        {
            if (IsConnected && !NeedsRenewal())
                return;

            lock (_lock)
            {
                if (_isLoggingIn)
                {
                    System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Login já em andamento — ignorando nova tentativa.");
                    return;
                }
                _isLoggingIn = true;
            }

            try
            {
                bool success = await LoginAsync(cancellationToken).ConfigureAwait(false);
                if (!success)
                    LogAndStatusBarError("Falha ao autenticar na API Gateway.");
            }
            finally
            {
                lock (_lock)
                {
                    _isLoggingIn = false;
                }
            }
        }

        /// <summary>
        /// Efetua o logout e limpa a sessão atual.
        /// </summary>
        public static async Task LogoutAsync(CancellationToken cancellationToken = default)
        {
            if (_client == null) return;

            try
            {
                await _client.PostAsync("logout", null, cancellationToken).ConfigureAwait(false);
                System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Logout enviado para a API Gateway.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Erro no logout (ignorado): {ex.Message}");
            }

            ClearSession();
        }

        /// <summary>
        /// Descarta completamente o cliente HTTP e recursos.
        /// </summary>
        public static void Dispose()
        {
            lock (_lock)
            {
                if (_isDisposed) return;

                ClearSessionInternal();
                _isDisposed = true;
            }

            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Recursos liberados (Dispose).");
        }

        #endregion

        #region Controle de sessão

        private static bool NeedsRenewal()
        {
            return _sessionExpiry == DateTime.MinValue ||
                   DateTime.Now.AddMinutes(1) >= _sessionExpiry ||
                   _client == null;
        }

        private static void ClearSession()
        {
            lock (_lock)
            {
                ClearSessionInternal();
            }
        }

        private static void ClearSessionInternal()
        {
            try
            {
                _client?.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Erro ao descartar HttpClient: {ex.Message}");
            }

            _client = null;
            _sessionExpiry = DateTime.MinValue;
            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão limpa.");
        }

        #endregion

        #region Login e autenticação

        private static async Task<bool> LoginAsync(CancellationToken cancellationToken = default)
        {
            ClearSessionInternal();

            // 🔐 Dados do app.config
            string apiUrl = ConfigurationManager.AppSettings["ApiGatewayUrl"];
            string user = ConfigurationManager.AppSettings["SAP_Username"];
            string pass = ConfigurationManager.AppSettings["SAP_Password"];
            string companyDB = Application.SBO_Application.Company.DatabaseName;

            if (string.IsNullOrEmpty(apiUrl) || string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass) || string.IsNullOrEmpty(companyDB))
            {
                LogAndStatusBarError("Configurações ApiGatewayUrl/User/Password/CompanyDB ausentes ou inválidas.");
                return false;
            }

            if (apiUrl.EndsWith("/"))
                apiUrl = apiUrl.TrimEnd('/');

            _baseUri = new Uri(apiUrl + "/");

            var payload = new { CompanyDB = companyDB, UserName = user, Password = pass };
            string jsonPayload = JsonSerializer.Serialize(payload);
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");


            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            var handler = new HttpClientHandler
            {
                CookieContainer = _cookies,
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
            };

            var client = new HttpClient(handler)
            {
                BaseAddress = _baseUri,
                Timeout = TimeSpan.FromMinutes(5)

            };
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36");
            System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] POST -> {client.BaseAddress}login | Payload: {jsonPayload}");

            try
            {
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                var response = await client.PostAsync("login", content, cancellationToken).ConfigureAwait(false);
                stopwatch.Stop();

                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Response ({(int)response.StatusCode}) {response.StatusCode} | Tempo: {stopwatch.ElapsedMilliseconds}ms | Body: {body}");

                if (!response.IsSuccessStatusCode)
                {
                    LogAndStatusBarError($"Falha no login da API Gateway ({response.StatusCode}): {body}");
                    client.Dispose();
                    return false;
                }

                int timeoutMinutes = 30;
                try
                {
                    using (JsonDocument document = JsonDocument.Parse(body))
                    {
                        if (document.RootElement.TryGetProperty("SessionTimeout", out JsonElement timeoutElement) &&
                            timeoutElement.TryGetInt32(out int parsed))
                        {
                            timeoutMinutes = parsed;
                        }
                    }
                }
                catch { }

                _sessionExpiry = DateTime.Now.AddMinutes(Math.Max(1, timeoutMinutes - 1));

                lock (_lock)
                {
                    _client = client;
                }

                System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] ✅ Login bem-sucedido. Sessão expira em: {_sessionExpiry}");
                return true;
            }
            catch (TaskCanceledException ex)
            {
                LogAndStatusBarError($"Timeout ou cancelamento no login da API Gateway: {ex.Message}");
                client.Dispose();
                return false;
            }
            catch (Exception ex)
            {
                LogAndStatusBarError($"Erro durante o login da API Gateway: {ex.Message}");
                client.Dispose();
                return false;
            }
        }

        #endregion

        #region Método CallAsync (padrão ServiceLayer)

        /// <summary>
        /// Executa uma requisição autenticada na API Gateway.
        /// Se a sessão estiver expirada, renova automaticamente e reenvia a chamada.
        /// </summary>
        public static async Task<HttpResponseMessage> CallAsync(HttpRequestMessage request, CancellationToken cancellationToken = default)
        {
            if (_isDisposed)
                throw new ObjectDisposedException(nameof(ApiGatewayConnector));

            await EnsureSessionAsync(cancellationToken).ConfigureAwait(false);

            if (_client == null)
                throw new InvalidOperationException("Cliente HTTP não inicializado.");

            try
            {
                System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Enviando requisição: {request.Method} {request.RequestUri}");

                var response = await _client.SendAsync(request, cancellationToken).ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                {
                    string error = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Erro HTTP {response.StatusCode}: {error}");

                    // Se sessão expirada → limpa e força novo login
                    if (response.StatusCode == HttpStatusCode.Unauthorized || error.Contains("Invalid session"))
                    {
                        System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão inválida detectada — renovando...");
                        ClearSession();
                        await EnsureSessionAsync(cancellationToken).ConfigureAwait(false);

                        // 🔁 Reenvia após renovar sessão
                        request.Headers.Remove("Cookie");
                        response = await _client.SendAsync(request, cancellationToken).ConfigureAwait(false);
                    }
                }

                return response;
            }
            catch (Exception ex)
            {
                LogAndStatusBarError($"Erro na chamada HTTP: {ex.Message}");
                throw;
            }
        }

        #endregion

        #region Helper

        private static void LogAndStatusBarError(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] ❌ {message}");
            try
            {
                Application.SBO_Application?.StatusBar?.SetText(
                    $"Erro API Gateway: {message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            catch { }
        }

        #endregion
    }
}
