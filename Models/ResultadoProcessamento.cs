using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItTech.Tool.AddonNFS.Models
{
    /// <summary>
    /// Resultado do processamento de uma linha
    /// </summary>
    public class ResultadoProcessamento
    {
        public string CodigoLinha { get; set; }
        public int NumeroLinha { get; set; }
        public bool Sucesso { get; set; }
        public int? DocEntry { get; set; }
        public int? DocNum { get; set; }
        public string Mensagem { get; set; }
        public DateTime DataProcessamento { get; set; }
        public bool PdfGeradoComSucesso { get; set; } = false;
        public string PdfMensagemErro { get; set; }
        public string PdfCaminhoCompleto { get; set; }

        public ResultadoProcessamento()
        {
            DataProcessamento = DateTime.Now;
        }
    }

}
