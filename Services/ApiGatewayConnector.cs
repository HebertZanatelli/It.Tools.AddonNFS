using System;
using System.Configuration;
using SAPbouiCOM;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using SAPbouiCOM.Framework; // Para Application.SBO_Application

namespace ItTech.Tool.AddonNFS.Services
{
    /// <summary>
    /// Gerencia a conexão e autenticação com o SAP Business One API Gateway.
    /// Utiliza o endpoint /login e gerencia a sessão via Cookies.
    /// Classe estática para manter uma única instância do HttpClient e estado de sessão.
    /// </summary>
    public static class ApiGatewayConnector // Usando static para simplicidade singleton
    {
        #region Campos Estáticos Privados

        private static HttpClient _apiGatewayClient = null;
        // CookieContainer é crucial para que o HttpClient gerencie automaticamente os cookies de sessão
        private static readonly CookieContainer _cookieContainer = new CookieContainer();
        private static Uri _apiGatewayBaseUri = null;
        private static DateTime _apiGatewayExpiry = DateTime.MinValue;
        private static readonly object _lockApiGateway = new object(); // Lock para thread-safety
        private static bool _isDisposed = false; // Controle para Dispose
        private static SemaphoreSlim _loginSemaphore = new SemaphoreSlim(1, 1); // Garante que apenas um login ocorra por vez

        #endregion

        #region Propriedades Públicas Estáticas

        /// <summary>
        /// Obtém a instância do HttpClient configurada e autenticada (se a sessão estiver ativa).
        /// Retorna null se não estiver conectado. Chame EnsureApiGatewaySessionAsync antes de usar.
        /// </summary>
        public static HttpClient Client
        {
            get
            {
                lock (_lockApiGateway)
                {
                    // Retorna o cliente apenas se a sessão for considerada válida e não descartado
                    return IsConnected ? _apiGatewayClient : null;
                }
            }
        }

        /// <summary>
        /// Verifica se a conexão com a API Gateway está ativa e a sessão não expirou.
        /// </summary>
        public static bool IsConnected
        {
            get
            {
                lock (_lockApiGateway)
                {
                    // Verifica se não foi descartado, se o cliente existe e se a sessão é válida
                    return !_isDisposed && _apiGatewayClient != null && DateTime.Now < _apiGatewayExpiry;
                }
            }
        }

        #endregion

        #region Métodos Públicos Estáticos

        /// <summary>
        /// Garante que existe uma sessão ativa com a API Gateway.
        /// Se a sessão não existir ou tiver expirado, tenta realizar o login de forma segura (evitando logins simultâneos).
        /// Lança uma exceção se o login falhar.
        /// </summary>
        /// <param name="cancellationToken">Token de cancelamento.</param>
        /// <exception cref="Exception">Lançada se o login falhar.</exception>
        /// <exception cref="ObjectDisposedException">Lançada se o conector já foi descartado.</exception>
        public static async Task EnsureApiGatewaySessionAsync(CancellationToken cancellationToken = default)
        {
            // Verificação rápida fora do lock (otimização)
            if (IsConnected) return;

            // Entra na seção crítica para verificar novamente e tentar o login se necessário
            await _loginSemaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                // Verifica novamente dentro do lock do semáforo, caso outra thread já tenha conectado
                if (IsConnected) return;

                lock (_lockApiGateway)
                {
                    if (_isDisposed) throw new ObjectDisposedException(nameof(ApiGatewayConnector));
                }

                System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão API Gateway inválida. Tentando login...");
                bool loginSuccess = await LoginApiGatewayAsyncInternal(cancellationToken).ConfigureAwait(false); // Chama método interno

                if (!loginSuccess)
                {
                    // LoginApiGatewayAsyncInternal já logou o erro específico.
                    throw new Exception("Falha ao estabelecer ou renovar a sessão com a API Gateway.");
                }
                System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão estabelecida/renovada.");
            }
            finally
            {
                _loginSemaphore.Release(); // Libera o semáforo
            }
        }


        /// <summary>
        /// Realiza o logout da API Gateway (best-effort) e limpa a sessão local.
        /// </summary>
        public static async Task LogoutAsync(CancellationToken cancellationToken = default)
        {
            HttpClient clientToLogout = null;
            lock (_lockApiGateway)
            {
                clientToLogout = _apiGatewayClient; // Pega a referência atual para usar fora do lock
            }

            if (clientToLogout != null)
            {
                try
                {
                    // Tenta chamar o endpoint /logout
                    // Ignoramos erros aqui, o principal é limpar localmente
                    await clientToLogout.PostAsync("logout", null, cancellationToken).ConfigureAwait(false);
                    System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Chamada /logout enviada (resultado ignorado).");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Exceção durante /logout (ignorado): {ex.Message}");
                }
            }

            ClearSession(); // Limpa o estado local
            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Sessão local e HttpClient limpos.");
        }


        /// <summary>
        /// Libera os recursos (HttpClient). Deve ser chamado ao descarregar o Add-on.
        /// </summary>
        public static void Dispose()
        {
            lock (_lockApiGateway) // Garante exclusividade ao descartar
            {
                if (_isDisposed) return; // Já descartado

                ClearSessionInternal(); // Limpa o cliente HTTP e estado
                _isDisposed = true; // Marca como descartado
            }
            _loginSemaphore.Dispose(); // Libera o semáforo
            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Recursos liberados.");
        }

        #endregion

        #region Métodos Privados Estáticos

        /// <summary>
        /// Lógica interna de login. NÃO chamar diretamente, use EnsureApiGatewaySessionAsync.
        /// Assume que já está dentro de um lock ou semáforo apropriado.
        /// Retorna true se o login for bem-sucedido, false caso contrário.
        /// </summary>
        private static async Task<bool> LoginApiGatewayAsyncInternal(CancellationToken cancellationToken = default)
        {
            // Limpa estado anterior antes de tentar novo login (dentro do lock/semáforo)
            ClearSessionInternal();

            string apiUrl = null;
            string user = null;
            string pass = null;
            string companyDB = null;

            try
            {
                apiUrl = ConfigurationManager.AppSettings["ApiGatewayUrl"];
                user = ConfigurationManager.AppSettings["SAP_Username"];
                pass = ConfigurationManager.AppSettings["SAP_Password"];

                companyDB = SAPbouiCOM.Framework.Application.SBO_Application.Company.DatabaseName; // DB da conexão atual do Addon

                if (string.IsNullOrEmpty(apiUrl) || string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass) || string.IsNullOrEmpty(companyDB))
                {
                    LogAndStatusBarError("Configurações ApiGatewayUrl/User/Password/CompanyDB ausentes ou inválidas.");
                    return false;
                }

                // Prepara URL base
                if (apiUrl.EndsWith("/")) apiUrl = apiUrl.TrimEnd('/');
                _apiGatewayBaseUri = new Uri(apiUrl + "/"); // Adiciona / para BaseAddress do HttpClient

                // Prepara Payload
                var loginPayload = new { CompanyDB = companyDB, UserName = user, Password = pass };
                string jsonPayload = JsonSerializer.Serialize(loginPayload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                // Prepara HttpClientHandler (NOVO a cada login para garantir estado limpo de cookies)
                HttpClientHandler handler = new HttpClientHandler
                {
                    CookieContainer = _cookieContainer, // Usa o container estático
                    ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true // SSL bypass
                };

                // Cria HttpClient (NOVO a cada login)
                var client = new HttpClient(handler) { BaseAddress = _apiGatewayBaseUri };
                client.Timeout = TimeSpan.FromMinutes(2);

                // --- Fazer a Chamada de Login ---
                HttpResponseMessage response = await client.PostAsync("login", content, cancellationToken).ConfigureAwait(false); // Caminho relativo

                if (response.IsSuccessStatusCode)
                {
                    // Login OK. Cookies foram armazenados no _cookieContainer.
                    var responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    int timeoutMinutes = 30; // Default
                    try
                    {
                        using (JsonDocument document = JsonDocument.Parse(responseBody))
                        {
                            if (document.RootElement.TryGetProperty("SessionTimeout", out JsonElement timeoutElement) && timeoutElement.TryGetInt32(out int parsedTimeout))
                            {
                                timeoutMinutes = parsedTimeout;
                            }
                        }
                    }
                    catch { /* Ignora erro no parse do timeout, usa default */ }

                    // Define expiração com margem de segurança
                    _apiGatewayExpiry = DateTime.Now.AddMinutes(Math.Max(1, timeoutMinutes - 1));

                    // Armazena o cliente configurado SOMENTE DENTRO DO LOCK
                    lock (_lockApiGateway)
                    {
                        _apiGatewayClient = client; // Guarda a nova instância do cliente
                    }

                    System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] Login bem-sucedido. Sessão expira em: {_apiGatewayExpiry}");
                    return true;
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    LogAndStatusBarError($"Falha no login da API Gateway ({response.StatusCode}): {error}");
                    client.Dispose(); // Descarta o cliente criado se o login falhou
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogAndStatusBarError($"Exceção durante login na API Gateway: {ex.Message}");
                // Garante que o cliente seja descartado em caso de exceção
                lock (_lockApiGateway)
                {
                    _apiGatewayClient?.Dispose();
                    _apiGatewayClient = null;
                }
                _apiGatewayExpiry = DateTime.MinValue;
                return false;
            }
        }


        /// <summary>
        /// Limpa o estado da sessão e descarta o HttpClient. Deve ser chamado dentro de um lock.
        /// </summary>
        private static void ClearSessionInternal()
        {
            _apiGatewayClient?.Dispose(); // Descarta o cliente HTTP se existir
            _apiGatewayClient = null;
            _apiGatewayExpiry = DateTime.MinValue;
            System.Diagnostics.Debug.WriteLine("[ApiGatewayConnector] Estado interno da sessão limpo.");
        }

        /// <summary>
        /// Método auxiliar para limpar a sessão de forma segura (usado por Logout e Dispose).
        /// </summary>
        private static void ClearSession()
        {
            lock (_lockApiGateway)
            {
                ClearSessionInternal();
            }
        }


        /// <summary>
        /// Helper para logar erro no Debug e mostrar na StatusBar do B1.
        /// </summary>
        private static void LogAndStatusBarError(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[ApiGatewayConnector] ERRO: {message}");
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application?.StatusBar?.SetText($"Erro API Gateway: {message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            catch { }
        }

        #endregion
    }
}
