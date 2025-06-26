using System;
using SAPbobsCOM;

namespace ItTech.Tool.AddonNFS.Database
{
    public class DatabaseSetup
    {
        private Company _company;

        public DatabaseSetup(Company company)
        {
            _company = company;
        }

        /// <summary>
        /// Cria todas as estruturas necessárias no banco de dados
        /// </summary>
        public void CreateDatabaseStructure()
        {
            try
            {
                // Criar tabelas de usuário
                CreateUserTables();

                // Criar campos de usuário
                CreateUserFields();

                // Criar objetos de usuário (UDO)
                // CreateUserObjects();
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao criar estrutura do banco: {ex.Message}");
            }
        }

        public void DeletarTabelaPorDI(string nomeTabela)
        {
            try
            {
                // IMPORTANTE: Nome SEM o @
                string tableName = nomeTabela.Replace("@", "");

                UserTablesMD oUserTablesMD = (UserTablesMD)_company.GetBusinessObject(BoObjectTypes.oUserTables);

                if (oUserTablesMD.GetByKey(tableName))
                {
                    int result = oUserTablesMD.Remove();
                    if (result != 0)
                    {
                        Console.WriteLine($"Erro: {_company.GetLastErrorDescription()}");
                    }
                    else
                    {
                        Console.WriteLine($"Tabela {tableName} removida com sucesso!");
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex.Message}");
            }
        }

        private void CreateUserTables()
        {


            // Tabela de Grupos de Lote
            CreateTable("IT_GRUPO_LOTE", "Grupos de Lote NFS-e", BoUTBTableType.bott_NoObjectAutoIncrement);

            // Tabela de Linhas do Lote
            CreateTable("IT_LINHA_LOTE", "Linhas de Lote NFS-e", BoUTBTableType.bott_NoObjectAutoIncrement);

            // Tabela de Log de Processamento
            CreateTable("ITCON_LOG_PROC", "Log de Processamento NFS-e", BoUTBTableType.bott_NoObjectAutoIncrement);
        }

        private void CreateUserFields()
        {
            // Campos da tabela IT_GRUPO_LOTE
            CreateField("@IT_GRUPO_LOTE", "Nome", "Nome do Grupo", BoFieldTypes.db_Alpha, 100);
            CreateField("@IT_GRUPO_LOTE", "DataLancamento", "Data de Lançamento", BoFieldTypes.db_Date);
            CreateField("@IT_GRUPO_LOTE", "DataDocumento", "Data do Documento", BoFieldTypes.db_Date);
            CreateField("@IT_GRUPO_LOTE", "NomeArquivo", "Nome do Arquivo", BoFieldTypes.db_Alpha, 254);
            CreateField("@IT_GRUPO_LOTE", "CaminhoArquivo", "Caminho do Arquivo", BoFieldTypes.db_Memo);
            CreateField("@IT_GRUPO_LOTE", "Status", "Status", BoFieldTypes.db_Alpha, 1);
            CreateField("@IT_GRUPO_LOTE", "TotalLinhas", "Total de Linhas", BoFieldTypes.db_Numeric);
            CreateField("@IT_GRUPO_LOTE", "LinhasProcessadas", "Linhas Processadas", BoFieldTypes.db_Numeric);
            CreateField("@IT_GRUPO_LOTE", "LinhasErro", "Linhas com Erro", BoFieldTypes.db_Numeric);
            CreateField("@IT_GRUPO_LOTE", "DataCriacao", "Data de Criação", BoFieldTypes.db_Date);
            CreateField("@IT_GRUPO_LOTE", "HoraCriacao", "Hora de Criação", BoFieldTypes.db_Date, 0, BoFldSubTypes.st_Time);

            // Campos da tabela IT_LINHA_LOTE
            CreateField("@IT_LINHA_LOTE", "GrupoCode", "Código do Grupo", BoFieldTypes.db_Alpha, 50);
            CreateField("@IT_LINHA_LOTE", "NumeroLinha", "Número da Linha", BoFieldTypes.db_Numeric);
            CreateField("@IT_LINHA_LOTE", "Filial", "Filial", BoFieldTypes.db_Alpha, 10);  // NOVO
            CreateField("@IT_LINHA_LOTE", "CodCliente", "Código do Cliente", BoFieldTypes.db_Alpha, 15);
            CreateField("@IT_LINHA_LOTE", "NomeCliente", "Nome do Cliente", BoFieldTypes.db_Alpha, 100);
            CreateField("@IT_LINHA_LOTE", "CodItem", "Código do Item", BoFieldTypes.db_Alpha, 20);
            CreateField("@IT_LINHA_LOTE", "DescItem", "Descrição do Item", BoFieldTypes.db_Alpha, 100);
            CreateField("@IT_LINHA_LOTE", "Utilizacao", "Utilização", BoFieldTypes.db_Alpha, 10);
            CreateField("@IT_LINHA_LOTE", "CodImposto", "Código do Imposto", BoFieldTypes.db_Alpha, 10);
            CreateField("@IT_LINHA_LOTE", "SeqNF", "Sequência NF", BoFieldTypes.db_Alpha, 10);
            CreateField("@IT_LINHA_LOTE", "CondPagto", "Condição de Pagamento", BoFieldTypes.db_Numeric);
            CreateField("@IT_LINHA_LOTE", "Valor", "Valor", BoFieldTypes.db_Float, 0, BoFldSubTypes.st_Price);
            CreateField("@IT_LINHA_LOTE", "ObsNF", "Observação NF", BoFieldTypes.db_Memo);  // NOVO
            CreateField("@IT_LINHA_LOTE", "TipoTrib", "Tipo Tributação", BoFieldTypes.db_Alpha, 50);  // NOVO
            CreateField("@IT_LINHA_LOTE", "DocEntry", "DocEntry NF", BoFieldTypes.db_Numeric);
            CreateField("@IT_LINHA_LOTE", "DocNum", "DocNum NF", BoFieldTypes.db_Numeric);
            CreateField("@IT_LINHA_LOTE", "Status", "Status", BoFieldTypes.db_Alpha, 1);
            CreateField("@IT_LINHA_LOTE", "MsgErro", "Mensagem de Erro", BoFieldTypes.db_Memo);
            CreateField("@IT_LINHA_LOTE", "Reprocessar", "Reprocessar", BoFieldTypes.db_Alpha, 1);
            CreateField("@IT_LINHA_LOTE", "DataProc", "Data Processamento", BoFieldTypes.db_Date);
            CreateField("@IT_LINHA_LOTE", "HoraProc", "Hora Processamento", BoFieldTypes.db_Date, 0, BoFldSubTypes.st_Time);

            // Valores válidos para campos de Status
            AddValidValues("@IT_GRUPO_LOTE", "Status", new string[,] {
                {"N", "Novo"},
                {"P", "Em Processamento"},
                {"C", "Processado Completo"},
                {"E", "Processado com Erros"},
                {"F", "Falha"}
            });

            AddValidValues("@IT_LINHA_LOTE", "Status", new string[,] {
                {"P", "Pendente"},
                {"R", "Processando"},
                {"S", "Sucesso"},
                {"E", "Erro"},
                {"I", "Ignorada"}
            });

            AddValidValues("@IT_LINHA_LOTE", "Reprocessar", new string[,] {
                {"Y", "Sim"},
                {"N", "Não"}
            });
        }

        private void CreateUserObjects()
        {
            // Criar UDO para Grupo de Lote
            CreateUDO(
           "ITTECH_GRPLOTE",                          // Nome correto do UDO para usar no GeneralService
           "Grupo de Lote NFS-e",                     // Nome visível
           "ITCON_GRUPO_LOTE",                        // Nome da tabela sem o "@"
           new string[] { "IT_LINHA_LOTE" },       // Tabelas filhas (sem "@")
           BoUDOObjType.boud_Document                 // Tipo Documento, pois tem linhas
       );
        }

        #region Métodos auxiliares

        private void CreateTable(string tableName, string tableDescription, BoUTBTableType tableType)
        {
            UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = (UserTablesMD)_company.GetBusinessObject(BoObjectTypes.oUserTables);

                if (!oUserTablesMD.GetByKey(tableName))
                {
                    oUserTablesMD.TableName = tableName;
                    oUserTablesMD.TableDescription = tableDescription;
                    oUserTablesMD.TableType = tableType;

                    if (oUserTablesMD.Add() != 0)
                    {
                        throw new Exception($"Erro ao criar tabela {tableName}: {_company.GetLastErrorDescription()}");
                    }
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
            }
        }

        private void CreateField(string tableName, string fieldName, string fieldDescription,
                                BoFieldTypes fieldType, int size = 0, BoFldSubTypes subType = BoFldSubTypes.st_None)
        {
            UserFieldsMD oUserFieldsMD = null;
            try
            {
                oUserFieldsMD = (UserFieldsMD)_company.GetBusinessObject(BoObjectTypes.oUserFields);

                oUserFieldsMD.TableName = tableName;
                oUserFieldsMD.Name = fieldName;
                oUserFieldsMD.Description = fieldDescription;
                oUserFieldsMD.Type = fieldType;

                if (size > 0) oUserFieldsMD.Size = size;
                if (subType != BoFldSubTypes.st_None) oUserFieldsMD.SubType = subType;

                if (oUserFieldsMD.Add() != 0)
                {
                    int errCode = _company.GetLastErrorCode();
                    if (errCode != -2035) // Campo já existe
                    {
                        throw new Exception($"Erro ao criar campo {fieldName}: {_company.GetLastErrorDescription()}");
                    }
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            }
        }

        private void AddValidValues(string tableName, string fieldName, string[,] validValues)
        {
            UserFieldsMD oUserFieldsMD = null;
            try
            {
                oUserFieldsMD = (UserFieldsMD)_company.GetBusinessObject(BoObjectTypes.oUserFields);

                if (oUserFieldsMD.GetByKey(oUserFieldsMD.TableName, oUserFieldsMD.FieldID))
                {
                    for (int i = 0; i < validValues.GetLength(0); i++)
                    {
                        oUserFieldsMD.ValidValues.Value = validValues[i, 0];
                        oUserFieldsMD.ValidValues.Description = validValues[i, 1];
                        oUserFieldsMD.ValidValues.Add();
                    }

                    if (oUserFieldsMD.Update() != 0)
                    {
                        throw new Exception($"Erro ao adicionar valores válidos: {_company.GetLastErrorDescription()}");
                    }
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            }
        }

        private void CreateUDO(string code, string name, string tableName, string[] childTables, BoUDOObjType objectType)
        {
            UserObjectsMD oUserObjectsMD = null;
            try
            {
                oUserObjectsMD = (UserObjectsMD)_company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

                if (!oUserObjectsMD.GetByKey(code))
                {
                    oUserObjectsMD.Code = code;
                    oUserObjectsMD.Name = name;
                    oUserObjectsMD.ObjectType = objectType;
                    oUserObjectsMD.TableName = tableName;
                    oUserObjectsMD.CanFind = BoYesNoEnum.tYES;
                    oUserObjectsMD.CanDelete = BoYesNoEnum.tYES;
                    oUserObjectsMD.CanCancel = BoYesNoEnum.tNO;
                    oUserObjectsMD.CanClose = BoYesNoEnum.tNO;

                    // Adicionar tabelas filhas
                    if (childTables != null)
                    {
                        foreach (string childTable in childTables)
                        {
                            oUserObjectsMD.ChildTables.TableName = childTable;
                            oUserObjectsMD.ChildTables.Add();
                        }
                    }

                    if (oUserObjectsMD.Add() != 0)
                    {
                        throw new Exception($"Erro ao criar UDO {code}: {_company.GetLastErrorDescription()}");
                    }
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);
            }
        }

        #endregion
    }
}