using System;
using System.Collections.Generic;
using System.Linq;
using ItTech.Tool.AddonNFS.Models;

namespace ItTech.Tool.AddonNFS.Utils
{
    /// <summary>
    /// Resultado da importação de dados do Excel
    /// </summary>
    public class ImportResult
    {
        public ImportResult()
        {
            Linhas = new List<LinhaImportacao>();
            Validacao = new Models.ValidacaoImportacao();
            Sucesso = false;
            DataImportacao = DateTime.Now;
        }

        /// <summary>
        /// Indica se a importação foi bem-sucedida
        /// </summary>
        public bool Sucesso { get; set; }

        /// <summary>
        /// Lista de linhas importadas
        /// </summary>
        public List<LinhaImportacao> Linhas { get; set; }

        /// <summary>
        /// Resultado da validação dos dados importados
        /// </summary>
        public Models.ValidacaoImportacao Validacao { get; set; }

        /// <summary>
        /// Mensagem de erro geral (se houver)
        /// </summary>
        public string MensagemErro { get; set; }

        /// <summary>
        /// Data e hora da importação
        /// </summary>
        public DateTime DataImportacao { get; set; }

        /// <summary>
        /// Número total de linhas processadas
        /// </summary>
        public int TotalLinhasProcessadas { get; set; }

        /// <summary>
        /// Número de linhas válidas
        /// </summary>
        public int LinhasValidas { get; set; }

        /// <summary>
        /// Número de linhas com erro
        /// </summary>
        public int LinhasComErro { get; set; }

        /// <summary>
        /// Caminho do arquivo importado
        /// </summary>
        public string CaminhoArquivo { get; set; }

        /// <summary>
        /// Nome do arquivo importado
        /// </summary>
        public string NomeArquivo { get; set; }

        /// <summary>
        /// Tamanho do arquivo em bytes
        /// </summary>
        public long TamanhoArquivo { get; set; }

        /// <summary>
        /// Adiciona uma linha ao resultado
        /// </summary>
        public void AdicionarLinha(LinhaImportacao linha)
        {
            if (linha == null) return;

            Linhas.Add(linha);
            TotalLinhasProcessadas++;

            if (linha.Status == StatusLinha.Pendente)
                LinhasValidas++;
            else if (linha.Status == StatusLinha.Erro)
                LinhasComErro++;
        }

        /// <summary>
        /// Adiciona um erro de validação
        /// </summary>
        public void AdicionarErro(string erro)
        {
            Validacao.AdicionarErro(erro);
            Sucesso = false;
        }

        /// <summary>
        /// Adiciona um aviso de validação
        /// </summary>
        public void AdicionarAviso(string aviso)
        {
            Validacao.AdicionarAviso(aviso);
        }

        /// <summary>
        /// Finaliza o resultado da importação
        /// </summary>
        public void Finalizar()
        {
            // Atualizar contadores finais
            LinhasValidas = Linhas.Count(l => l.Status == StatusLinha.Pendente);
            LinhasComErro = Linhas.Count(l => l.Status == StatusLinha.Erro);
            TotalLinhasProcessadas = Linhas.Count;

            // Determinar sucesso geral
            Sucesso = Validacao.Valida && LinhasValidas > 0;

            // Se não há linhas válidas, considerar como erro
            if (LinhasValidas == 0 && TotalLinhasProcessadas > 0)
            {
                AdicionarErro("Nenhuma linha válida foi encontrada no arquivo");
            }
        }

        /// <summary>
        /// Obtém resumo da importação
        /// </summary>
        public string ObterResumo()
        {
            return $"Importação {(Sucesso ? "bem-sucedida" : "com erros")}: " +
                   $"{TotalLinhasProcessadas} linhas processadas, " +
                   $"{LinhasValidas} válidas, " +
                   $"{LinhasComErro} com erro";
        }

        /// <summary>
        /// Obtém estatísticas detalhadas
        /// </summary>
        public Dictionary<string, object> ObterEstatisticas()
        {
            return new Dictionary<string, object>
            {
                ["Sucesso"] = Sucesso,
                ["TotalLinhas"] = TotalLinhasProcessadas,
                ["LinhasValidas"] = LinhasValidas,
                ["LinhasComErro"] = LinhasComErro,
                ["PercentualSucesso"] = TotalLinhasProcessadas > 0 ?
                    Math.Round((double)LinhasValidas / TotalLinhasProcessadas * 100, 2) : 0,
                ["DataImportacao"] = DataImportacao,
                ["NomeArquivo"] = NomeArquivo,
                ["TamanhoArquivo"] = TamanhoArquivo,
                ["TempoProcessamento"] = DateTime.Now - DataImportacao
            };
        }
    }
}