using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Utils;

namespace ItTech.Tool.AddonNFS.Utils
{
    /// <summary>
    /// Serviço para importação de dados do Excel
    /// </summary>
    public class ExcelImportService
    {
        private readonly Dictionary<string, int> _colunasPadrao;

        public ExcelImportService()
        {
            // Mapeamento padrão das colunas do Excel
            _colunasPadrao = new Dictionary<string, int>
            {
                ["NumeroLinha"] = 0,
                ["CodigoCliente"] = 1,
                ["NomeCliente"] = 2,
                ["CodigoItem"] = 3,
                ["DescricaoItem"] = 4,
                ["Valor"] = 5,
                ["CondicaoPagamento"] = 6,
                ["CodigoImposto"] = 7,
                ["Utilizacao"] = 8,
                ["SeqNF"] = 9
            };
        }

        /// <summary>
        /// Importa dados de um arquivo Excel
        /// </summary>
        public ImportResult ImportarExcel(string caminhoArquivo)
        {
            var resultado = new ImportResult
            {
                CaminhoArquivo = caminhoArquivo,
                NomeArquivo = Path.GetFileName(caminhoArquivo)
            };

            try
            {
                if (!ValidarArquivo(caminhoArquivo, resultado))
                    return resultado;

                // Para esta implementação básica, vamos simular a leitura do Excel
                // Em uma implementação real, você usaria EPPlus, ClosedXML ou similar
                var linhas = LerArquivoExcel(caminhoArquivo);

                ProcessarLinhas(linhas, resultado);

                resultado.Finalizar();
            }
            catch (Exception ex)
            {
                resultado.AdicionarErro($"Erro ao importar arquivo: {ex.Message}");
            }

            return resultado;
        }

        /// <summary>
        /// Valida se o arquivo pode ser importado
        /// </summary>
        private bool ValidarArquivo(string caminhoArquivo, ImportResult resultado)
        {
            if (string.IsNullOrEmpty(caminhoArquivo))
            {
                resultado.AdicionarErro("Caminho do arquivo não informado");
                return false;
            }

            if (!File.Exists(caminhoArquivo))
            {
                resultado.AdicionarErro("Arquivo não encontrado");
                return false;
            }

            var extensao = Path.GetExtension(caminhoArquivo).ToLower();
            if (extensao != ".xlsx" && extensao != ".xls" && extensao != ".csv")
            {
                resultado.AdicionarErro("Formato de arquivo não suportado. Use .xlsx, .xls ou .csv");
                return false;
            }

            var info = new FileInfo(caminhoArquivo);
            resultado.TamanhoArquivo = info.Length;

            // Verificar se o arquivo não está muito grande (limite de 10MB)
            if (info.Length > 10 * 1024 * 1024)
            {
                resultado.AdicionarAviso("Arquivo muito grande. O processamento pode ser lento.");
            }

            return true;
        }

        /// <summary>
        /// Lê o arquivo Excel (implementação básica)
        /// </summary>
        private List<string[]> LerArquivoExcel(string caminhoArquivo)
        {
            var linhas = new List<string[]>();

            try
            {
                var extensao = Path.GetExtension(caminhoArquivo).ToLower();

                if (extensao == ".csv")
                {
                    // Leitura de CSV
                    var linhasTexto = File.ReadAllLines(caminhoArquivo);
                    foreach (var linha in linhasTexto.Skip(1)) // Pular cabeçalho
                    {
                        var colunas = linha.Split(',', ';');
                        linhas.Add(colunas);
                    }
                }
                else
                {
                    // Para arquivos Excel (.xlsx, .xls), seria necessário usar uma biblioteca
                    // Como EPPlus ou ClosedXML. Por enquanto, vamos simular dados
                    linhas.AddRange(GerarDadosExemplo());
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao ler arquivo: {ex.Message}");
            }

            return linhas;
        }

        /// <summary>
        /// Gera dados de exemplo para teste (remover em produção)
        /// </summary>
        private List<string[]> GerarDadosExemplo()
        {
            return new List<string[]>
            {
                new[] { "1", "C001", "Cliente Teste 1", "I001", "Item Teste 1", "100.00", "1", "ICMS", "13", "46" },
                new[] { "2", "C002", "Cliente Teste 2", "I002", "Item Teste 2", "200.00", "2", "ICMS", "13", "46" },
                new[] { "3", "C003", "Cliente Teste 3", "I003", "Item Teste 3", "300.00", "1", "ICMS", "13", "46" }
            };
        }

        /// <summary>
        /// Processa as linhas lidas do Excel
        /// </summary>
        private void ProcessarLinhas(List<string[]> linhasExcel, ImportResult resultado)
        {
            for (int i = 0; i < linhasExcel.Count; i++)
            {
                try
                {
                    var linhaExcel = linhasExcel[i];
                    var linha = ConverterLinhaExcel(linhaExcel, i + 1);

                    if (linha != null)
                    {
                        ValidarLinha(linha);
                        resultado.AdicionarLinha(linha);
                    }
                }
                catch (Exception ex)
                {
                    var linhaErro = new LinhaImportacao
                    {
                        NumeroLinha = i + 1,
                        Status = StatusLinha.Erro,
                        MensagemErro = $"Erro ao processar linha {i + 1}: {ex.Message}"
                    };
                    resultado.AdicionarLinha(linhaErro);
                }
            }
        }

        /// <summary>
        /// Converte uma linha do Excel em LinhaImportacao
        /// </summary>
        private LinhaImportacao ConverterLinhaExcel(string[] colunas, int numeroLinha)
        {
            if (colunas == null || colunas.Length == 0)
                return null;

            var linha = new LinhaImportacao
            {
                NumeroLinha = numeroLinha,
                Status = StatusLinha.Pendente,
                Code = Guid.NewGuid().ToString()
            };

            try
            {
                // Mapear colunas usando o dicionário padrão
                linha.CodigoCliente = ObterValorColuna(colunas, "CodigoCliente");
                linha.NomeCliente = ObterValorColuna(colunas, "NomeCliente");
                linha.CodigoItem = ObterValorColuna(colunas, "CodigoItem");
                linha.DescricaoItem = ObterValorColuna(colunas, "DescricaoItem");
                linha.CondicaoPagamento = ObterValorColuna(colunas, "CondicaoPagamento");
                linha.CodigoImposto = ObterValorColuna(colunas, "CodigoImposto");
                linha.Utilizacao = ObterValorColuna(colunas, "Utilizacao");
                linha.CodSeq = ObterValorColuna(colunas, "SeqNF");

                // Converter valor
                var valorStr = ObterValorColuna(colunas, "Valor");
                if (decimal.TryParse(valorStr?.Replace(",", "."), out decimal valor))
                {
                    linha.Valor = valor;
                }
                else
                {
                    linha.Valor = 0;
                    linha.Status = StatusLinha.Erro;
                    linha.MensagemErro = "Valor inválido";
                }
            }
            catch (Exception ex)
            {
                linha.Status = StatusLinha.Erro;
                linha.MensagemErro = $"Erro ao converter dados: {ex.Message}";
            }

            return linha;
        }

        /// <summary>
        /// Obtém valor de uma coluna específica
        /// </summary>
        private string ObterValorColuna(string[] colunas, string nomeColuna)
        {
            if (!_colunasPadrao.ContainsKey(nomeColuna))
                return null;

            var indice = _colunasPadrao[nomeColuna];
            if (indice >= 0 && indice < colunas.Length)
            {
                return colunas[indice]?.Trim();
            }

            return null;
        }

        /// <summary>
        /// Valida uma linha importada
        /// </summary>
        private void ValidarLinha(LinhaImportacao linha)
        {
            var erros = new List<string>();

            // Validações obrigatórias
            if (string.IsNullOrEmpty(linha.CodigoCliente))
                erros.Add("Código do cliente é obrigatório");

            if (string.IsNullOrEmpty(linha.CodigoItem))
                erros.Add("Código do item é obrigatório");

            if (linha.Valor <= 0)
                erros.Add("Valor deve ser maior que zero");

            // Validações de formato
            if (!string.IsNullOrEmpty(linha.CodigoCliente) && linha.CodigoCliente.Length > 20)
                erros.Add("Código do cliente muito longo (máx. 20 caracteres)");

            if (!string.IsNullOrEmpty(linha.CodigoItem) && linha.CodigoItem.Length > 20)
                erros.Add("Código do item muito longo (máx. 20 caracteres)");

            if (!string.IsNullOrEmpty(linha.NomeCliente) && linha.NomeCliente.Length > 100)
                erros.Add("Nome do cliente muito longo (máx. 100 caracteres)");

            // Se há erros, marcar linha como erro
            if (erros.Count > 0)
            {
                linha.Status = StatusLinha.Erro;
                linha.MensagemErro = string.Join("; ", erros);
            }
        }

        /// <summary>
        /// Configura mapeamento personalizado de colunas
        /// </summary>
        public void ConfigurarMapeamentoColunas(Dictionary<string, int> mapeamento)
        {
            foreach (var item in mapeamento)
            {
                if (_colunasPadrao.ContainsKey(item.Key))
                {
                    _colunasPadrao[item.Key] = item.Value;
                }
            }
        }

        /// <summary>
        /// Obtém template de exemplo para Excel
        /// </summary>
        public string[] ObterCabecalhoTemplate()
        {
            return new[]
            {
                "Numero Linha",
                "Codigo Cliente",
                "Nome Cliente",
                "Codigo Item",
                "Descricao Item",
                "Valor",
                "Condicao Pagamento",
                "Codigo Imposto",
                "Utilizacao",
                "Sequencia NF"
            };
        }

        /// <summary>
        /// Valida se o arquivo tem o formato esperado
        /// </summary>
        public bool ValidarFormatoArquivo(string caminhoArquivo)
        {
            try
            {
                var linhas = LerArquivoExcel(caminhoArquivo);

                // Verificar se tem pelo menos uma linha de dados
                if (linhas.Count == 0)
                    return false;

                // Verificar se a primeira linha tem o número mínimo de colunas
                var primeiraLinha = linhas.FirstOrDefault();
                return primeiraLinha != null && primeiraLinha.Length >= 6; // Mínimo: cliente, item, valor, etc.
            }
            catch
            {
                return false;
            }
        }
    }
}