using System;

namespace ItTech.Tool.AddonNFS.Configuration
{
    /// <summary>
    /// Classe para gerenciar configurações do Service Layer
    /// Versão simplificada sem dependência do System.Configuration
    /// </summary>
    public static class ServiceLayerConfig
    {
        // Configurações padrão - podem ser alteradas via código ou arquivo de configuração externo
        private static string _baseUrl = "https://hanab1:50000/b1s/v2";
        private static string _username = "";
        private static string _password = "";
        private static string _companyDB = "";
        private static int _timeoutSeconds = 120;

        /// <summary>
        /// URL base do Service Layer
        /// </summary>
        public static string BaseUrl
        {
            get { return _baseUrl; }
            set { _baseUrl = value; }
        }

        /// <summary>
        /// Usuário para autenticação no Service Layer
        /// </summary>
        public static string Username
        {
            get { return _username; }
            set { _username = value; }
        }

        /// <summary>
        /// Senha para autenticação no Service Layer
        /// </summary>
        public static string Password
        {
            get { return _password; }
            set { _password = value; }
        }

        /// <summary>
        /// Database da empresa no SAP
        /// </summary>
        public static string CompanyDB
        {
            get { return _companyDB; }
            set { _companyDB = value; }
        }

        /// <summary>
        /// Timeout para requisições HTTP (em segundos)
        /// </summary>
        public static int TimeoutSeconds
        {
            get { return _timeoutSeconds; }
            set { _timeoutSeconds = value; }
        }

        /// <summary>
        /// Configura todas as propriedades de uma vez
        /// </summary>
        public static void Configure(string baseUrl, string username, string password, string companyDB, int timeoutSeconds = 120)
        {
            BaseUrl = baseUrl;
            Username = username;
            Password = password;
            CompanyDB = companyDB;
            TimeoutSeconds = timeoutSeconds;
        }

        /// <summary>
        /// Verifica se todas as configurações necessárias estão presentes
        /// </summary>
        public static bool IsConfigured()
        {
            return !string.IsNullOrEmpty(BaseUrl) &&
                   !string.IsNullOrEmpty(Username) &&
                   !string.IsNullOrEmpty(Password) &&
                   !string.IsNullOrEmpty(CompanyDB);
        }

        /// <summary>
        /// Obtém mensagem de erro para configurações faltantes
        /// </summary>
        public static string GetConfigurationErrorMessage()
        {
            var missing = new System.Collections.Generic.List<string>();

            if (string.IsNullOrEmpty(BaseUrl))
                missing.Add("BaseUrl");
            if (string.IsNullOrEmpty(Username))
                missing.Add("Username");
            if (string.IsNullOrEmpty(Password))
                missing.Add("Password");
            if (string.IsNullOrEmpty(CompanyDB))
                missing.Add("CompanyDB");

            if (missing.Count > 0)
            {
                return $"Configurações faltantes: {string.Join(", ", missing)}";
            }

            return string.Empty;
        }

        /// <summary>
        /// Carrega configurações de um arquivo de texto simples (opcional)
        /// Formato: chave=valor (uma por linha)
        /// </summary>
        public static void LoadFromFile(string filePath)
        {
            try
            {
                if (System.IO.File.Exists(filePath))
                {
                    var lines = System.IO.File.ReadAllLines(filePath);
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                            continue;

                        var parts = line.Split('=');
                        if (parts.Length == 2)
                        {
                            var key = parts[0].Trim();
                            var value = parts[1].Trim();

                            switch (key.ToLower())
                            {
                                case "baseurl":
                                    BaseUrl = value;
                                    break;
                                case "username":
                                    Username = value;
                                    break;
                                case "password":
                                    Password = value;
                                    break;
                                case "companydb":
                                    CompanyDB = value;
                                    break;
                                case "timeoutseconds":
                                    if (int.TryParse(value, out int timeout))
                                        TimeoutSeconds = timeout;
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao carregar configurações do arquivo: {ex.Message}");
            }
        }

        /// <summary>
        /// Salva configurações em um arquivo de texto simples
        /// </summary>
        public static void SaveToFile(string filePath)
        {
            try
            {
                var lines = new[]
                {
                    "# Configurações do Service Layer",
                    $"BaseUrl={BaseUrl}",
                    $"Username={Username}",
                    $"Password={Password}",
                    $"CompanyDB={CompanyDB}",
                    $"TimeoutSeconds={TimeoutSeconds}"
                };

                System.IO.File.WriteAllLines(filePath, lines);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar configurações no arquivo: {ex.Message}");
            }
        }
    }
}

