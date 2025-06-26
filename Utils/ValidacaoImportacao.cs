using System;
using System.Collections.Generic;
using System.Linq;

namespace ItTech.Tool.AddonNFS.Utils
{
    /// <summary>
    /// Resultado da validação de importação
    /// </summary>
    public class ValidacaoImportacao
    {
        public ValidacaoImportacao()
        {
            Erros = new List<string>();
            Avisos = new List<string>();
            Valida = true;
        }

        /// <summary>
        /// Indica se a validação passou sem erros críticos
        /// </summary>
        public bool Valida { get; set; }

        /// <summary>
        /// Lista de erros encontrados
        /// </summary>
        public List<string> Erros { get; set; }

        /// <summary>
        /// Lista de avisos encontrados
        /// </summary>
        public List<string> Avisos { get; set; }

        /// <summary>
        /// Adiciona um erro à validação
        /// </summary>
        public void AdicionarErro(string erro)
        {
            if (!string.IsNullOrEmpty(erro))
            {
                Erros.Add(erro);
                Valida = false;
            }
        }

        /// <summary>
        /// Adiciona um aviso à validação
        /// </summary>
        public void AdicionarAviso(string aviso)
        {
            if (!string.IsNullOrEmpty(aviso))
            {
                Avisos.Add(aviso);
            }
        }

        /// <summary>
        /// Adiciona múltiplos erros
        /// </summary>
        public void AdicionarErros(IEnumerable<string> erros)
        {
            if (erros != null)
            {
                foreach (var erro in erros.Where(e => !string.IsNullOrEmpty(e)))
                {
                    AdicionarErro(erro);
                }
            }
        }

        /// <summary>
        /// Adiciona múltiplos avisos
        /// </summary>
        public void AdicionarAvisos(IEnumerable<string> avisos)
        {
            if (avisos != null)
            {
                foreach (var aviso in avisos.Where(a => !string.IsNullOrEmpty(a)))
                {
                    AdicionarAviso(aviso);
                }
            }
        }

        /// <summary>
        /// Limpa todos os erros e avisos
        /// </summary>
        public void Limpar()
        {
            Erros.Clear();
            Avisos.Clear();
            Valida = true;
        }

        /// <summary>
        /// Obtém resumo da validação
        /// </summary>
        public string ObterResumo()
        {
            if (Valida && Avisos.Count == 0)
                return "Validação passou sem problemas";

            var resumo = new List<string>();

            if (!Valida)
                resumo.Add($"{Erros.Count} erro(s)");

            if (Avisos.Count > 0)
                resumo.Add($"{Avisos.Count} aviso(s)");

            return string.Join(", ", resumo);
        }

        /// <summary>
        /// Obtém todos os problemas (erros + avisos) formatados
        /// </summary>
        public List<string> ObterTodosProblemas()
        {
            var problemas = new List<string>();

            foreach (var erro in Erros)
            {
                problemas.Add($"ERRO: {erro}");
            }

            foreach (var aviso in Avisos)
            {
                problemas.Add($"AVISO: {aviso}");
            }

            return problemas;
        }

        /// <summary>
        /// Verifica se há problemas (erros ou avisos)
        /// </summary>
        public bool TemProblemas()
        {
            return Erros.Count > 0 || Avisos.Count > 0;
        }

        /// <summary>
        /// Obtém contagem total de problemas
        /// </summary>
        public int TotalProblemas()
        {
            return Erros.Count + Avisos.Count;
        }

        /// <summary>
        /// Combina esta validação com outra
        /// </summary>
        public void Combinar(ValidacaoImportacao outraValidacao)
        {
            if (outraValidacao != null)
            {
                AdicionarErros(outraValidacao.Erros);
                AdicionarAvisos(outraValidacao.Avisos);
            }
        }

        /// <summary>
        /// Cria uma cópia da validação
        /// </summary>
        public ValidacaoImportacao Clonar()
        {
            var clone = new ValidacaoImportacao
            {
                Valida = this.Valida
            };

            clone.Erros.AddRange(this.Erros);
            clone.Avisos.AddRange(this.Avisos);

            return clone;
        }

        /// <summary>
        /// Converte para string com formatação
        /// </summary>
        public override string ToString()
        {
            if (!TemProblemas())
                return "Validação OK";

            var resultado = new List<string>();

            if (Erros.Count > 0)
            {
                resultado.Add($"Erros ({Erros.Count}):");
                resultado.AddRange(Erros.Select(e => $"  - {e}"));
            }

            if (Avisos.Count > 0)
            {
                resultado.Add($"Avisos ({Avisos.Count}):");
                resultado.AddRange(Avisos.Select(a => $"  - {a}"));
            }

            return string.Join(Environment.NewLine, resultado);
        }
    }
}

