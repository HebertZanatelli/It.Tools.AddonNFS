using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ItTech.Tool.AddonNFS.Utils
{
    /// <summary>
    /// Utilitários gerais para o add-on
    /// </summary>
    public static class UtilsHelper
    {
        /// <summary>
        /// Valida se uma string é um código válido (alfanumérico)
        /// </summary>
        public static bool IsValidCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return false;

            return code.All(c => char.IsLetterOrDigit(c) || c == '_' || c == '-');
        }

        /// <summary>
        /// Gera um código único baseado em timestamp
        /// </summary>
        public static string GenerateUniqueCode(string prefix = "")
        {
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var random = new Random().Next(100, 999);

            if (string.IsNullOrEmpty(prefix))
                return $"{timestamp}{random}";

            return $"{prefix}_{timestamp}{random}";
        }

        /// <summary>
        /// Formata um valor decimal para exibição
        /// </summary>
        public static string FormatCurrency(decimal value)
        {
            return value.ToString("C2");
        }

        /// <summary>
        /// Formata uma data para exibição
        /// </summary>
        public static string FormatDate(DateTime date)
        {
            return date.ToString("dd/MM/yyyy");
        }

        /// <summary>
        /// Formata uma data e hora para exibição
        /// </summary>
        public static string FormatDateTime(DateTime dateTime)
        {
            return dateTime.ToString("dd/MM/yyyy HH:mm:ss");
        }

        /// <summary>
        /// Trunca uma string se ela exceder o tamanho máximo
        /// </summary>
        public static string TruncateString(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;

            return text.Substring(0, maxLength - 3) + "...";
        }

        /// <summary>
        /// Remove caracteres especiais de uma string
        /// </summary>
        public static string RemoveSpecialCharacters(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return new string(text.Where(c => char.IsLetterOrDigit(c) || char.IsWhiteSpace(c)).ToArray());
        }

        /// <summary>
        /// Converte uma string para decimal de forma segura
        /// </summary>
        public static decimal SafeParseDecimal(string value, decimal defaultValue = 0)
        {
            if (string.IsNullOrWhiteSpace(value))
                return defaultValue;

            // Tentar diferentes formatos
            value = value.Replace(",", ".");

            if (decimal.TryParse(value, out decimal result))
                return result;

            return defaultValue;
        }

        /// <summary>
        /// Converte uma string para inteiro de forma segura
        /// </summary>
        public static int SafeParseInt(string value, int defaultValue = 0)
        {
            if (string.IsNullOrWhiteSpace(value))
                return defaultValue;

            if (int.TryParse(value, out int result))
                return result;

            return defaultValue;
        }

        /// <summary>
        /// Converte uma string para DateTime de forma segura
        /// </summary>
        public static DateTime SafeParseDateTime(string value, DateTime defaultValue = default)
        {
            if (string.IsNullOrWhiteSpace(value))
                return defaultValue == default ? DateTime.Now : defaultValue;

            if (DateTime.TryParse(value, out DateTime result))
                return result;

            return defaultValue == default ? DateTime.Now : defaultValue;
        }

        /// <summary>
        /// Verifica se um arquivo existe e é acessível
        /// </summary>
        public static bool IsFileAccessible(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return false;

                // Tentar abrir o arquivo para verificar se está acessível
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Obtém o tamanho de um arquivo em formato legível
        /// </summary>
        public static string GetFileSize(string filePath)
        {
            try
            {
                var info = new FileInfo(filePath);
                long bytes = info.Length;

                string[] sizes = { "B", "KB", "MB", "GB" };
                double len = bytes;
                int order = 0;

                while (len >= 1024 && order < sizes.Length - 1)
                {
                    order++;
                    len = len / 1024;
                }

                return $"{len:0.##} {sizes[order]}";
            }
            catch
            {
                return "Desconhecido";
            }
        }

        /// <summary>
        /// Cria um backup de um arquivo
        /// </summary>
        public static string CreateBackup(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return null;

                var directory = Path.GetDirectoryName(filePath);
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var extension = Path.GetExtension(filePath);
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                var backupPath = Path.Combine(directory, $"{fileName}_backup_{timestamp}{extension}");
                File.Copy(filePath, backupPath);

                return backupPath;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Limpa arquivos temporários antigos
        /// </summary>
        public static void CleanupTempFiles(string directory, int daysOld = 7)
        {
            try
            {
                if (!Directory.Exists(directory))
                    return;

                var cutoffDate = DateTime.Now.AddDays(-daysOld);
                var files = Directory.GetFiles(directory, "*temp*", SearchOption.TopDirectoryOnly);

                foreach (var file in files)
                {
                    var info = new FileInfo(file);
                    if (info.CreationTime < cutoffDate)
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch
                        {
                            // Ignorar erros ao deletar arquivos individuais
                        }
                    }
                }
            }
            catch
            {
                // Ignorar erros na limpeza
            }
        }

        /// <summary>
        /// Valida se uma string contém apenas números
        /// </summary>
        public static bool IsNumeric(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            return value.All(char.IsDigit);
        }

        /// <summary>
        /// Capitaliza a primeira letra de cada palavra
        /// </summary>
        public static string ToTitleCase(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            var words = text.Split(' ');
            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > 0)
                {
                    words[i] = char.ToUpper(words[i][0]) +
                              (words[i].Length > 1 ? words[i].Substring(1).ToLower() : "");
                }
            }

            return string.Join(" ", words);
        }

        /// <summary>
        /// Obtém uma lista de extensões de arquivo suportadas
        /// </summary>
        public static List<string> GetSupportedFileExtensions()
        {
            return new List<string> { ".xlsx", ".xls", ".csv", ".txt" };
        }

        /// <summary>
        /// Verifica se uma extensão de arquivo é suportada
        /// </summary>
        public static bool IsSupportedFileExtension(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            var extension = Path.GetExtension(filePath).ToLower();
            return GetSupportedFileExtensions().Contains(extension);
        }

        /// <summary>
        /// Gera um hash simples para uma string
        /// </summary>
        public static string GenerateSimpleHash(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            int hash = 0;
            foreach (char c in input)
            {
                hash = ((hash << 5) - hash) + c;
                hash = hash & hash; // Convert to 32-bit integer
            }

            return Math.Abs(hash).ToString("X8");
        }

        /// <summary>
        /// Converte bytes para string legível
        /// </summary>
        public static string BytesToString(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            double len = bytes;
            int order = 0;

            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }

            return $"{len:0.##} {sizes[order]}";
        }
    }
}

