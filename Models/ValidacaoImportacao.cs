using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItTech.Tool.AddonNFS.Models
{
    /// <summary>
    /// Classe para validação dos dados importados
    /// </summary>
    public class ValidacaoImportacao
    {
        public bool Valida { get; set; }
        public List<string> Erros { get; set; }
        public List<string> Avisos { get; set; }

        public ValidacaoImportacao()
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
