using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItTech.Tool.AddonNFS.Models
{
    /// <summary>
    /// Representa um grupo de lote para processamento de NFS-e
    /// </summary>
    public class GrupoLote
    {
        // Identificação
        public string Code { get; set; }
        public string Nome { get; set; }

        // Datas
        public DateTime DataLancamento { get; set; }
        public DateTime DataDocumento { get; set; }
        public DateTime DataCriacao { get; set; }
        public DateTime? DataUltimoProcessamento { get; set; }

        // Arquivo
        public string NomeArquivo { get; set; }
        public string CaminhoArquivo { get; set; }
        public string CaminhoPDF { get; set; }


        public string TipoDocumento { get; set; }

        // Status e Controle
        public StatusGrupo Status { get; set; }
        public int TotalLinhas { get; set; }
        public int LinhasProcessadas { get; set; }
        public int LinhasErro { get; set; }

        // Usuário
        public string UsuarioCriacao { get; set; }

        // Linhas do grupo
        public List<LinhaImportacao> Linhas { get; set; }

        // Propriedades calculadas
        public int LinhasPendentes
        {
            get { return TotalLinhas - LinhasProcessadas - LinhasErro; }
        }

        public decimal PercentualProcessado
        {
            get
            {
                if (TotalLinhas == 0) return 0;
                return (decimal)(LinhasProcessadas + LinhasErro) * 100 / TotalLinhas;
            }
        }

        public bool PodeProcessar
        {
            get
            {
                return Status == StatusGrupo.Novo ||
                       Status == StatusGrupo.ProcessadoParcial ||
                       (Status == StatusGrupo.Erro && LinhasErro > 0);
            }
        }

        public bool ProcessamentoCompleto
        {
            get { return Status == StatusGrupo.ProcessadoCompleto; }
        }

        public GrupoLote()
        {
            Linhas = new List<LinhaImportacao>();
            Status = StatusGrupo.Novo;
            DataCriacao = DateTime.Now;
            TotalLinhas = 0;
            LinhasProcessadas = 0;
            LinhasErro = 0;
        }

        /// <summary>
        /// Atualiza os contadores baseado nas linhas
        /// </summary>
        public void AtualizarContadores()
        {
            if (Linhas != null && Linhas.Count > 0)
            {
                TotalLinhas = Linhas.Count;
                LinhasProcessadas = Linhas.Count(l => l.Status == StatusLinha.Sucesso);
                LinhasErro = Linhas.Count(l => l.Status == StatusLinha.Erro);
            }
        }

        /// <summary>
        /// Retorna uma descrição amigável do status
        /// </summary>
        public string DescricaoStatus()
        {
            switch (Status)
            {
                case StatusGrupo.Novo:
                    return "Novo - Aguardando processamento";
                case StatusGrupo.EmProcessamento:
                    return "Em processamento...";
                case StatusGrupo.ProcessadoCompleto:
                    return $"Processado com sucesso ({LinhasProcessadas} notas criadas)";
                case StatusGrupo.ProcessadoParcial:
                    return $"Processado parcialmente ({LinhasProcessadas} sucesso, {LinhasErro} erros)";
                case StatusGrupo.Erro:
                    return $"Erro no processamento ({LinhasErro} erros)";
                default:
                    return "Status desconhecido";
            }
        }

        /// <summary>
        /// Valida se o grupo está pronto para processamento
        /// </summary>
        public ValidacaoGrupo ValidarParaProcessamento()
        {
            var validacao = new ValidacaoGrupo();

            if (string.IsNullOrWhiteSpace(Nome))
                validacao.AdicionarErro("Nome do grupo não informado");

            if (DataLancamento == DateTime.MinValue)
                validacao.AdicionarErro("Data de lançamento não informada");

            if (DataDocumento == DateTime.MinValue)
                validacao.AdicionarErro("Data do documento não informada");

            if (Linhas == null || Linhas.Count == 0)
                validacao.AdicionarErro("Nenhuma linha foi importada para este grupo");

            if (Status == StatusGrupo.ProcessadoCompleto)
                validacao.AdicionarAviso("Este grupo já foi processado completamente");

            if (Status == StatusGrupo.EmProcessamento)
                validacao.AdicionarErro("Este grupo está em processamento por outro usuário");

            return validacao;
        }
    }

    /// <summary>
    /// Classe auxiliar para validação do grupo
    /// </summary>
    public class ValidacaoGrupo
    {
        public bool Valida { get; set; }
        public List<string> Erros { get; set; }
        public List<string> Avisos { get; set; }

        public ValidacaoGrupo()
        {
            Valida = true;
            Erros = new List<string>();
            Avisos = new List<string>();
        }

        public void AdicionarErro(string erro)
        {
            Valida = false;
            Erros.Add(erro);
        }

        public void AdicionarAviso(string aviso)
        {
            Avisos.Add(aviso);
        }
    }
}