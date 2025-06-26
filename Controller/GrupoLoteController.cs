using System;
using System.Collections.Generic;
using ItTech.Tool.AddonNFS.Models;
using ClosedXML.Excel; // Adicionar via NuGet
using SAPbobsCOM;

namespace ItTech.Tool.AddonNFS.Controllers
{
    /// <summary>
    /// Controller principal para gerenciar grupos de lote
    /// </summary>
    public class GrupoLoteController
    {
        private Company _company;
        private string _connectionString;

        public GrupoLoteController(Company company)
        {
            _company = company;
        }

        /// <summary>
        /// Cria um novo grupo de lote
        /// </summary>
        public string CriarGrupo(GrupoLote grupo)
        {
            try
            {
                // Gera um código único manualmente
                string codigo = GerarCodigoUnico();

                // A tabela do SAP deve ser acessada SEM o "@"
                UserTable userTable = (UserTable)_company.UserTables.Item("IT_GRUPO_LOTE");
                // userTable.Code = codigo;
                userTable.Name = grupo.Nome;
                userTable.UserFields.Fields.Item("U_Nome").Value = grupo.Nome;
                userTable.UserFields.Fields.Item("U_DataLancamento").Value = grupo.DataLancamento;
                userTable.UserFields.Fields.Item("U_DataDocumento").Value = grupo.DataDocumento;
                userTable.UserFields.Fields.Item("U_NomeArquivo").Value = grupo.NomeArquivo;
                userTable.UserFields.Fields.Item("U_CaminhoArquivo").Value = grupo.CaminhoArquivo;
                userTable.UserFields.Fields.Item("U_Status").Value = "N"; // Novo
                userTable.UserFields.Fields.Item("U_TotalLinhas").Value = 0;
                userTable.UserFields.Fields.Item("U_LinhasProcessadas").Value = 0;
                userTable.UserFields.Fields.Item("U_LinhasErro").Value = 0;

                userTable.UserFields.Fields.Item("U_DataCriacao").Value = grupo.DataLancamento; // ou DateTime.Now.Date;
                userTable.UserFields.Fields.Item("U_HoraCriacao").Value = grupo.DataLancamento; // ou DateTime.Now;

                int ret = userTable.Add();
                if (ret != 0)
                {
                    throw new Exception(_company.GetLastErrorDescription());
                }

                string codigoGerado = null;
                // Após userTable.Add();
                if (ret == 0)
                {
                    Recordset rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    string query = $@"SELECT ""Code"" FROM ""@IT_GRUPO_LOTE"" WHERE ""Name"" = '{grupo.Nome}'";
                    rs.DoQuery(query);
                    if (!rs.EoF)
                    {
                        codigoGerado = rs.Fields.Item("Code").Value.ToString();
                    }
                }
                else
                {
                    throw new Exception(_company.GetLastErrorDescription());
                }

                return codigoGerado;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao criar grupo: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ MÉTODO CORRIGIDO: Atualiza status do grupo usando os campos existentes
        /// </summary>
        public bool AtualizarStatusGrupoComTotais(string grupoCode, StatusGrupo status, int sucessos, int erros, int pendentes)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string statusChar = ObterCodigoStatus(status);
                int totalLinhas = sucessos + erros + pendentes;

                string query = $@"
                    UPDATE ""@IT_GRUPO_LOTE""
                    SET ""U_Status"" = '{statusChar}',
                        ""U_TotalLinhas"" = {totalLinhas},
                        ""U_LinhasProcessadas"" = {sucessos},
                        ""U_LinhasErro"" = {erros}
                    WHERE ""Code"" = '{grupoCode}'";

                oRecordset.DoQuery(query);

                // ✅ VERIFICAR se a atualização funcionou
                string verificarQuery = $@"SELECT ""Code"" FROM ""@IT_GRUPO_LOTE"" WHERE ""Code"" = '{grupoCode}'";
                oRecordset.DoQuery(verificarQuery);

                return !oRecordset.EoF; // Retorna true se encontrou o registro
            }
            catch (Exception ex)
            {
                // Log do erro para debug (opcional)
                System.Diagnostics.Debug.WriteLine($"Erro ao atualizar totais do grupo {grupoCode}: {ex.Message}");
                return false;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// ✅ MÉTODO CORRIGIDO: Atualiza apenas o status do grupo (versão simples)
        /// </summary>
        public bool AtualizarStatusGrupoSimples(string grupoCode, StatusGrupo status)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string statusChar = ObterCodigoStatus(status);

                string query = $@"
                    UPDATE ""@IT_GRUPO_LOTE""
                    SET ""U_Status"" = '{statusChar}'
                    WHERE ""Code"" = '{grupoCode}'";

                oRecordset.DoQuery(query);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// Lista todos os grupos existentes
        /// </summary>
        public List<GrupoLote> ListarGrupos()
        {
            List<GrupoLote> grupos = new List<GrupoLote>();
            Recordset oRecordset = null;

            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = @"
                    SELECT 
                        ""Code"",
                        ""U_Nome"",
                        ""U_DataLancamento"",
                        ""U_DataDocumento"",
                        ""U_NomeArquivo"",
                        ""U_Status"",
                        ""U_TotalLinhas"",
                        ""U_LinhasProcessadas"",
                        ""U_LinhasErro"",
                        ""U_DataCriacao""
                    FROM ""@IT_GRUPO_LOTE""
                    ORDER BY ""Code"" DESC";

                oRecordset.DoQuery(query);

                while (!oRecordset.EoF)
                {
                    // ✅ USAR os nomes corretos da classe GrupoLote
                    int totalLinhas = Convert.ToInt32(oRecordset.Fields.Item("U_TotalLinhas").Value ?? 0);
                    int linhasProcessadas = Convert.ToInt32(oRecordset.Fields.Item("U_LinhasProcessadas").Value ?? 0);
                    int linhasErro = Convert.ToInt32(oRecordset.Fields.Item("U_LinhasErro").Value ?? 0);

                    grupos.Add(new GrupoLote
                    {
                        Code = oRecordset.Fields.Item("Code").Value.ToString(),
                        Nome = oRecordset.Fields.Item("U_Nome").Value.ToString(),
                        DataLancamento = (DateTime)oRecordset.Fields.Item("U_DataLancamento").Value,
                        DataDocumento = (DateTime)oRecordset.Fields.Item("U_DataDocumento").Value,
                        NomeArquivo = oRecordset.Fields.Item("U_NomeArquivo").Value.ToString(),
                        Status = ConverterStatus(oRecordset.Fields.Item("U_Status").Value.ToString()),
                        // ✅ USAR os nomes corretos da classe
                        TotalLinhas = totalLinhas,
                        LinhasProcessadas = linhasProcessadas,
                        LinhasErro = linhasErro
                        // LinhasPendentes é calculado automaticamente pela propriedade
                    });
                    oRecordset.MoveNext();
                }
                return grupos;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// Obtém um grupo específico com suas linhas
        /// </summary>
        public GrupoLote ObterGrupo(string code)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Buscar dados do grupo
                string query = $@"
                    SELECT 
                        ""Code"",
                        ""U_Nome"",
                        ""U_DataLancamento"",
                        ""U_DataDocumento"",
                        ""U_NomeArquivo"",
                        ""U_CaminhoArquivo"",
                        ""U_Status"",
                        ""U_TotalLinhas"",
                        ""U_LinhasProcessadas"",
                        ""U_LinhasErro""
                    FROM ""@IT_GRUPO_LOTE""
                    WHERE ""Code"" = '{code}'";

                oRecordset.DoQuery(query);

                if (oRecordset.EoF)
                {
                    throw new Exception($"Grupo '{code}' não encontrado");
                }

                // ✅ USAR os nomes corretos da classe GrupoLote  
                int totalLinhas = Convert.ToInt32(oRecordset.Fields.Item("U_TotalLinhas").Value ?? 0);
                int linhasProcessadas = Convert.ToInt32(oRecordset.Fields.Item("U_LinhasProcessadas").Value ?? 0);
                int linhasErro = Convert.ToInt32(oRecordset.Fields.Item("U_LinhasErro").Value ?? 0);

                GrupoLote grupo = new GrupoLote
                {
                    Code = oRecordset.Fields.Item("Code").Value.ToString(),
                    Nome = oRecordset.Fields.Item("U_Nome").Value.ToString(),
                    DataLancamento = (DateTime)oRecordset.Fields.Item("U_DataLancamento").Value,
                    DataDocumento = (DateTime)oRecordset.Fields.Item("U_DataDocumento").Value,
                    NomeArquivo = oRecordset.Fields.Item("U_NomeArquivo").Value.ToString(),
                    CaminhoArquivo = oRecordset.Fields.Item("U_CaminhoArquivo").Value.ToString(),
                    Status = ConverterStatus(oRecordset.Fields.Item("U_Status").Value.ToString()),
                    // ✅ USAR os nomes corretos da classe
                    TotalLinhas = totalLinhas,
                    LinhasProcessadas = linhasProcessadas,
                    LinhasErro = linhasErro
                    // LinhasPendentes é calculado automaticamente pela propriedade
                };

                // Carregar linhas
                grupo.Linhas = ObterLinhasGrupo(code);

                return grupo;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao obter grupo: {ex.Message}");
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// Atualiza o status de um grupo (método legado mantido para compatibilidade)
        /// </summary>
        public void AtualizarStatusGrupo(string code, StatusGrupo status, int linhasProcessadas, int linhasErro)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string statusStr = ConverterStatusParaString(status);

                string updateQuery = $@"
                    UPDATE ""@IT_GRUPO_LOTE""
                    SET ""U_Status"" = '{statusStr}',
                        ""U_LinhasProcessadas"" = {linhasProcessadas},
                        ""U_LinhasErro"" = {linhasErro}
                    WHERE ""Code"" = '{code}'";

                oRecordset.DoQuery(updateQuery);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao atualizar status: {ex.Message}");
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// Atualiza o total de linhas de um grupo
        /// </summary>
        public void AtualizarTotalLinhas(string grupoCode, int totalLinhas)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = $@"
                    UPDATE ""@IT_GRUPO_LOTE""
                    SET ""U_TotalLinhas"" = {totalLinhas}
                    WHERE ""Code"" = '{grupoCode}'";

                oRecordset.DoQuery(query);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao atualizar total de linhas: {ex.Message}");
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// Obtém o status/etapa atual de um grupo
        /// </summary>
        public int ObterEtapaAtual(string code)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = $@"
                    SELECT ""U_Status"", ""U_TotalLinhas"", ""U_LinhasProcessadas""
                    FROM ""@IT_GRUPO_LOTE""
                    WHERE ""Code"" = '{code}'";

                oRecordset.DoQuery(query);

                if (!oRecordset.EoF)
                {
                    string status = oRecordset.Fields.Item("U_Status").Value.ToString();
                    int totalLinhas = Convert.ToInt32(oRecordset.Fields.Item("U_TotalLinhas").Value);

                    // Determinar etapa baseado no status
                    switch (status)
                    {
                        case "N": // Novo
                            return totalLinhas > 0 ? 2 : 1; // Se tem linhas, vai para etapa 2
                        case "P": // Em Processamento
                        case "C": // Processado Completo
                        case "E": // Processado com Erros
                            return 3; // Etapa de visualização de resultados
                        default:
                            return 1;
                    }
                }
                return 1;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// ✅ NOVO MÉTODO: Obtém estatísticas detalhadas de um grupo
        /// </summary>
        public EstatisticasGrupo ObterEstatisticasGrupo(string grupoCode)
        {
            Recordset oRecordset = null;
            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Contar linhas por status
                string query = $@"
                    SELECT 
                        COUNT(*) AS Total,
                        SUM(CASE WHEN ""U_Status"" = 'S' THEN 1 ELSE 0 END) AS Sucessos,
                        SUM(CASE WHEN ""U_Status"" = 'E' THEN 1 ELSE 0 END) AS Erros,
                        SUM(CASE WHEN ""U_Status"" = 'P' OR ""U_Status"" IS NULL THEN 1 ELSE 0 END) AS Pendentes
                    FROM ""@IT_LINHA_LOTE""
                    WHERE ""U_GrupoCode"" = '{grupoCode}'";

                oRecordset.DoQuery(query);

                if (!oRecordset.EoF)
                {
                    return new EstatisticasGrupo
                    {
                        TotalLinhas = Convert.ToInt32(oRecordset.Fields.Item("Total").Value ?? 0),
                        TotalSucessos = Convert.ToInt32(oRecordset.Fields.Item("Sucessos").Value ?? 0),
                        TotalErros = Convert.ToInt32(oRecordset.Fields.Item("Erros").Value ?? 0),
                        TotalPendentes = Convert.ToInt32(oRecordset.Fields.Item("Pendentes").Value ?? 0)
                    };
                }

                return new EstatisticasGrupo();
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao obter estatísticas do grupo: {ex.Message}");
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        #region Métodos auxiliares

        private string GerarCodigoUnico()
        {
            return $"GRP{DateTime.Now:yyyyMMddHHmmss}";
        }

        /// <summary>
        /// ✅ MÉTODO CORRIGIDO: Converte StatusGrupo para código de caractere único
        /// </summary>
        private string ObterCodigoStatus(StatusGrupo status)
        {
            switch (status)
            {
                case StatusGrupo.Novo: return "N";
                case StatusGrupo.EmProcessamento: return "P";
                case StatusGrupo.ProcessadoCompleto: return "C";
                case StatusGrupo.ProcessadoParcial: return "E";
                case StatusGrupo.Erro: return "F";
                default: return "N";
            }
        }

        private StatusGrupo ConverterStatus(string status)
        {
            switch (status)
            {
                case "N": return StatusGrupo.Novo;
                case "P": return StatusGrupo.EmProcessamento;
                case "C": return StatusGrupo.ProcessadoCompleto;
                case "E": return StatusGrupo.ProcessadoParcial;
                case "F": return StatusGrupo.Erro;
                default: return StatusGrupo.Novo;
            }
        }

        private string ConverterStatusParaString(StatusGrupo status)
        {
            switch (status)
            {
                case StatusGrupo.Novo: return "N";
                case StatusGrupo.EmProcessamento: return "P";
                case StatusGrupo.ProcessadoCompleto: return "C";
                case StatusGrupo.ProcessadoParcial: return "E";
                case StatusGrupo.Erro: return "F";
                default: return "N";
            }
        }

        public List<LinhaImportacao> ObterLinhasGrupo(string grupoCode)
        {
            List<LinhaImportacao> linhas = new List<LinhaImportacao>();
            Recordset oRecordset = null;

            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = $@"
                    SELECT * FROM ""@IT_LINHA_LOTE""
                    WHERE ""U_GrupoCode"" = '{grupoCode}'
                    ORDER BY ""U_NumeroLinha""";

                oRecordset.DoQuery(query);

                while (!oRecordset.EoF)
                {
                    linhas.Add(new LinhaImportacao
                    {
                        Code = oRecordset.Fields.Item("Code").Value.ToString(),
                        GrupoCode = grupoCode,
                        NumeroLinha = Convert.ToInt32(oRecordset.Fields.Item("U_NumeroLinha").Value),
                        Filial = oRecordset.Fields.Item("U_Filial").Value?.ToString(),  // NOVO
                        CodigoCliente = oRecordset.Fields.Item("U_CodCliente").Value.ToString(),
                        NomeCliente = oRecordset.Fields.Item("U_NomeCliente").Value.ToString(),
                        CodigoItem = oRecordset.Fields.Item("U_CodItem").Value.ToString(),
                        DescricaoItem = oRecordset.Fields.Item("U_DescItem").Value.ToString(),
                        Utilizacao = oRecordset.Fields.Item("U_Utilizacao").Value.ToString(),
                        CodigoImposto = oRecordset.Fields.Item("U_CodImposto").Value.ToString(),
                        CodSeq = oRecordset.Fields.Item("U_SeqNF").Value.ToString(),
                        CondicaoPagamento = oRecordset.Fields.Item("U_CondPagto").Value.ToString(),
                        Valor = Convert.ToDecimal(oRecordset.Fields.Item("U_Valor").Value),
                        ObservacaoNF = oRecordset.Fields.Item("U_ObsNF").Value?.ToString(),  // NOVO
                        TipoTributacao = oRecordset.Fields.Item("U_TipoTrib").Value?.ToString(),  // NOVO
                        Status = ConverterStatusLinha(oRecordset.Fields.Item("U_Status").Value?.ToString() ?? "P"),
                        DocEntry = oRecordset.Fields.Item("U_DocEntry").Value == null ? (int?)null : Convert.ToInt32(oRecordset.Fields.Item("U_DocEntry").Value),
                        DocNum = oRecordset.Fields.Item("U_DocNum").Value == null ? (int?)null : Convert.ToInt32(oRecordset.Fields.Item("U_DocNum").Value),
                        MensagemErro = oRecordset.Fields.Item("U_MsgErro").Value?.ToString(),
                        Reprocessar = oRecordset.Fields.Item("U_Reprocessar").Value?.ToString() == "Y"
                    });
                    oRecordset.MoveNext();
                }
                return linhas;
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        private StatusLinha ConverterStatusLinha(string status)
        {
            switch (status)
            {
                case "P": return StatusLinha.Pendente;
                case "R": return StatusLinha.Processando;
                case "S": return StatusLinha.Sucesso;
                case "E": return StatusLinha.Erro;
                case "I": return StatusLinha.Ignorada;
                default: return StatusLinha.Pendente;
            }
        }

        #endregion
    }

    /// <summary>
    /// ✅ CLASSE NOVA: Para retornar estatísticas detalhadas do grupo
    /// </summary>
    public class EstatisticasGrupo
    {
        public int TotalLinhas { get; set; }
        public int TotalSucessos { get; set; }
        public int TotalErros { get; set; }
        public int TotalPendentes { get; set; }
    }
}