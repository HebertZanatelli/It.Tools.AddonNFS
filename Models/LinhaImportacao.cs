using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItTech.Tool.AddonNFS.Models
{
    /// <summary>
    /// Representa uma linha importada do Excel
    /// </summary>
    public class LinhaImportacao
    {
        public string Code { get; set; }
        public string GrupoCode { get; set; }
        public int NumeroLinha { get; set; }

        // Dados do Excel
        public string Filial { get; set; }                    // NOVO - Col 1
        public string CodigoCliente { get; set; }             // Col 2
        public string NomeCliente { get; set; }               // Col 3
        public string CodigoItem { get; set; }                // Col 4
        public string DescricaoItem { get; set; }             // Col 5
        public string Utilizacao { get; set; }                // Col 6
        public string CodigoImposto { get; set; }             // Col 7 - MANTIDO
        public string CodSeq { get; set; }                    // Col 8 - RENOMEADO (era SeqNF)
        public string CondicaoPagamento { get; set; }         // Col 9
        public decimal Valor { get; set; }                    // Col 10
        public string ObservacaoNF { get; set; }              // Col 11 - NOVO
        public string TipoTributacao { get; set; }            // Col 12 - NOVO

        // Controle de processamento
        public bool Selecionada { get; set; }
        public StatusLinha Status { get; set; }
        public int? DocEntry { get; set; }
        public int? DocNum { get; set; }
        public string MensagemErro { get; set; }
        public bool Reprocessar { get; set; }
        public DateTime? DataProcessamento { get; set; }

        public LinhaImportacao()
        {
            Status = StatusLinha.Pendente;
            Selecionada = true;
            Reprocessar = false;
        }
    }

}
