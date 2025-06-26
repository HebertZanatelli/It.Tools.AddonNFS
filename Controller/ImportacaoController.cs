using ClosedXML.Excel;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Utils;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValidacaoImportacao = ItTech.Tool.AddonNFS.Models.ValidacaoImportacao;

namespace ItTech.Tool.AddonNFS.Controller
{

    /// <summary>
    /// Controller para importação de arquivos Excel
    /// </summary>
    public class ImportacaoController
    {
        private Company _company;

        public ImportacaoController(Company company)
        {
            _company = company;
        }

        /// <summary>
        /// Importa arquivo Excel e valida os dados
        /// </summary>
        public (List<LinhaImportacao> linhas, ValidacaoImportacao validacao) ImportarExcel(string caminhoArquivo)
        {
            List<LinhaImportacao> linhas = new List<LinhaImportacao>();
            ValidacaoImportacao validacao = new ValidacaoImportacao();

            try
            {
                using (var workbook = new XLWorkbook(caminhoArquivo))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed().Skip(1); // Pular cabeçalho

                    int numeroLinha = 1;

                    foreach (var row in rows)
                    {
                        try
                        {
                            var linha = new LinhaImportacao
                            {
                                NumeroLinha = numeroLinha,
                                Filial = row.Cell(1).GetString().Trim(),              // Col 1 - NOVO
                                CodigoCliente = row.Cell(2).GetString().Trim(),       // Col 2
                                NomeCliente = row.Cell(3).GetString().Trim(),         // Col 3
                                CodigoItem = row.Cell(4).GetString().Trim(),          // Col 4
                                DescricaoItem = row.Cell(5).GetString().Trim(),       // Col 5
                                Utilizacao = row.Cell(6).GetString().Trim(),          // Col 6
                                CodigoImposto = row.Cell(7).GetString().Trim(),       // Col 7 - MANTIDO
                                CodSeq = row.Cell(8).GetString().Trim(),              // Col 8 - RENOMEADO
                                CondicaoPagamento = row.Cell(9).GetString().Trim(),   // Col 9
                                Valor = row.Cell(10).GetValue<decimal>(),             // Col 10
                                ObservacaoNF = row.Cell(11).GetString().Trim(),       // Col 11 - NOVO
                                TipoTributacao = row.Cell(12).GetString().Trim()      // Col 12 - NOVO
                            };

                            // Validar linha
                            ValidarLinha(linha, numeroLinha, validacao);
                            linhas.Add(linha);
                        }
                        catch (Exception ex)
                        {
                            validacao.AdicionarErro($"Linha {numeroLinha}: Erro ao ler dados - {ex.Message}");
                        }
                        numeroLinha++;
                    }

                    if (linhas.Count == 0)
                    {
                        validacao.AdicionarErro("Nenhuma linha válida encontrada no arquivo");
                    }
                }

                return (linhas, validacao);
            }
            catch (Exception ex)
            {
                validacao.AdicionarErro($"Erro ao abrir arquivo: {ex.Message}");
                return (linhas, validacao);
            }
        }

        /// <summary>
        /// Salva as linhas importadas no banco
        /// </summary>
        public void SalvarLinhas(string grupoCode, List<LinhaImportacao> linhas)
        {
            UserTable userTable = null;
            Recordset oRecordset = null;

            try
            {
                // Opção 1: Usar UserTable (mais lento mas mais seguro)
                userTable = (UserTable)_company.UserTables.Item("IT_LINHA_LOTE");

                foreach (var linha in linhas)
                {
                    string code = $"{grupoCode}_{linha.NumeroLinha:D4}";

                    userTable.Name = code;
                    userTable.UserFields.Fields.Item("U_GrupoCode").Value = grupoCode;
                    userTable.UserFields.Fields.Item("U_NumeroLinha").Value = linha.NumeroLinha;
                    userTable.UserFields.Fields.Item("U_Filial").Value = linha.Filial ?? "";  // NOVO
                    userTable.UserFields.Fields.Item("U_CodCliente").Value = linha.CodigoCliente;
                    userTable.UserFields.Fields.Item("U_NomeCliente").Value = linha.NomeCliente;
                    userTable.UserFields.Fields.Item("U_CodItem").Value = linha.CodigoItem;
                    userTable.UserFields.Fields.Item("U_DescItem").Value = linha.DescricaoItem;
                    userTable.UserFields.Fields.Item("U_Utilizacao").Value = linha.Utilizacao ?? "";
                    userTable.UserFields.Fields.Item("U_CodImposto").Value = linha.CodigoImposto ?? "";
                    userTable.UserFields.Fields.Item("U_SeqNF").Value = linha.CodSeq ?? "";
                    userTable.UserFields.Fields.Item("U_CondPagto").Value = linha.CondicaoPagamento ?? "";
                    userTable.UserFields.Fields.Item("U_Valor").Value = Convert.ToDouble(linha.Valor);
                    userTable.UserFields.Fields.Item("U_ObsNF").Value = linha.ObservacaoNF ?? "";  // NOVO
                    userTable.UserFields.Fields.Item("U_TipoTrib").Value = linha.TipoTributacao ?? "";  // NOVO
                    userTable.UserFields.Fields.Item("U_Status").Value = "P"; // Pendente
                    userTable.UserFields.Fields.Item("U_Reprocessar").Value = "N";

                    int ret = userTable.Add();
                    if (ret != 0)
                    {
                        throw new Exception($"Erro ao salvar linha {linha.NumeroLinha}: {_company.GetLastErrorDescription()}");
                    }

                    linha.Code = code;
                }

                // Atualizar o total de linhas no grupo usando SQL
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string updateQuery = $@"
                        UPDATE ""@IT_GRUPO_LOTE"" 
                        SET ""U_TotalLinhas"" = {linhas.Count}
                        WHERE ""Code"" = '{grupoCode}'";

                oRecordset.DoQuery(updateQuery);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar linhas: {ex.Message}");
            }
            finally
            {
                if (userTable != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(userTable);
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        /// <summary>
        /// Importa arquivo Excel e valida os dados com callback de progresso
        /// </summary>
        public (List<LinhaImportacao> linhas, ValidacaoImportacao validacao) ImportarExcelComProgresso(
            string caminhoArquivo,
            Action<int, string> onProgress)
        {
            List<LinhaImportacao> linhas = new List<LinhaImportacao>();
            ValidacaoImportacao validacao = new ValidacaoImportacao();

            try
            {
                using (var workbook = new XLWorkbook(caminhoArquivo))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed().Skip(1); // Pular cabeçalho

                    // Contar total de linhas primeiro
                    int totalLinhas = rows.Count();
                    int linhasProcessadas = 0;
                    int ultimoPercentualReportado = 0;

                    onProgress?.Invoke(0, $"Lendo arquivo Excel... {totalLinhas} linhas encontradas");

                    int numeroLinha = 1;

                    foreach (var row in rows)
                    {
                        try
                        {
                            var linha = new LinhaImportacao
                            {
                                NumeroLinha = numeroLinha,
                                Filial = row.Cell(1).GetString().Trim(),
                                CodigoCliente = row.Cell(2).GetString().Trim(),
                                NomeCliente = row.Cell(3).GetString().Trim(),
                                CodigoItem = row.Cell(4).GetString().Trim(),
                                DescricaoItem = row.Cell(5).GetString().Trim(),
                                Utilizacao = row.Cell(6).GetString().Trim(),
                                CodigoImposto = row.Cell(7).GetString().Trim(),
                                CodSeq = row.Cell(8).GetString().Trim(),
                                CondicaoPagamento = row.Cell(9).GetString().Trim(),
                                Valor = row.Cell(10).GetValue<decimal>(),
                                ObservacaoNF = row.Cell(11).GetString().Trim(),
                                TipoTributacao = row.Cell(12).GetString().Trim()
                            };

                            // Validar linha
                            ValidarLinha(linha, numeroLinha, validacao);
                            linhas.Add(linha);
                        }
                        catch (Exception ex)
                        {
                            validacao.AdicionarErro($"Linha {numeroLinha}: Erro ao ler dados - {ex.Message}");
                        }

                        linhasProcessadas++;
                        int percentualAtual = (linhasProcessadas * 100) / totalLinhas;

                        // Reportar progresso de forma inteligente
                        bool deveReportarProgresso = false;

                        if (totalLinhas <= 100)
                        {
                            // Arquivos pequenos: a cada 10 linhas
                            deveReportarProgresso = linhasProcessadas % 10 == 0;
                        }
                        else if (totalLinhas <= 1000)
                        {
                            // Arquivos médios: a cada 5%
                            deveReportarProgresso = percentualAtual >= ultimoPercentualReportado + 5;
                        }
                        else
                        {
                            // Arquivos grandes: a cada 2%
                            deveReportarProgresso = percentualAtual >= ultimoPercentualReportado + 2;
                        }

                        // Sempre reportar a última linha
                        if (linhasProcessadas == totalLinhas)
                        {
                            deveReportarProgresso = true;
                        }

                        if (deveReportarProgresso)
                        {
                            onProgress?.Invoke(linhasProcessadas,
                                $"Lendo Excel... {linhasProcessadas}/{totalLinhas} linhas ({percentualAtual}%)");
                            ultimoPercentualReportado = percentualAtual;
                        }

                        numeroLinha++;
                    }

                    if (linhas.Count == 0)
                    {
                        validacao.AdicionarErro("Nenhuma linha válida encontrada no arquivo");
                    }
                    else
                    {
                        onProgress?.Invoke(totalLinhas, $"Leitura concluída! {linhas.Count} linhas válidas");
                    }
                }

                return (linhas, validacao);
            }
            catch (Exception ex)
            {
                validacao.AdicionarErro($"Erro ao abrir arquivo: {ex.Message}");
                return (linhas, validacao);
            }
        }

        // OPCIONAL: Versão otimizada para arquivos muito grandes
        /// <summary>
        /// Importa arquivo Excel em lotes para melhor performance em arquivos grandes
        /// </summary>
        public (List<LinhaImportacao> linhas, ValidacaoImportacao validacao) ImportarExcelEmLotesComProgresso(
            string caminhoArquivo,
            Action<int, string> onProgress,
            int tamanhoLote = 100)
        {
            List<LinhaImportacao> linhas = new List<LinhaImportacao>();
            ValidacaoImportacao validacao = new ValidacaoImportacao();

            try
            {
                using (var workbook = new XLWorkbook(caminhoArquivo))
                {
                    var worksheet = workbook.Worksheet(1);
                    var allRows = worksheet.RowsUsed().Skip(1).ToList(); // Pular cabeçalho

                    int totalLinhas = allRows.Count;
                    int linhasProcessadas = 0;

                    onProgress?.Invoke(0, $"Preparando leitura de {totalLinhas} linhas...");

                    // Processar em lotes
                    for (int i = 0; i < allRows.Count; i += tamanhoLote)
                    {
                        var lote = allRows.Skip(i).Take(tamanhoLote).ToList();

                        foreach (var row in lote)
                        {
                            int numeroLinha = i + allRows.IndexOf(row) + 2; // +2 por causa do cabeçalho e índice zero

                            try
                            {
                                var linha = new LinhaImportacao
                                {
                                    NumeroLinha = numeroLinha - 1,
                                    Filial = row.Cell(1).GetString().Trim(),
                                    CodigoCliente = row.Cell(2).GetString().Trim(),
                                    NomeCliente = row.Cell(3).GetString().Trim(),
                                    CodigoItem = row.Cell(4).GetString().Trim(),
                                    DescricaoItem = row.Cell(5).GetString().Trim(),
                                    Utilizacao = row.Cell(6).GetString().Trim(),
                                    CodigoImposto = row.Cell(7).GetString().Trim(),
                                    CodSeq = row.Cell(8).GetString().Trim(),
                                    CondicaoPagamento = row.Cell(9).GetString().Trim(),
                                    Valor = row.Cell(10).GetValue<decimal>(),
                                    ObservacaoNF = row.Cell(11).GetString().Trim(),
                                    TipoTributacao = row.Cell(12).GetString().Trim()
                                };

                                // Validar linha
                                ValidarLinha(linha, numeroLinha - 1, validacao);
                                linhas.Add(linha);
                            }
                            catch (Exception ex)
                            {
                                validacao.AdicionarErro($"Linha {numeroLinha}: Erro ao ler dados - {ex.Message}");
                            }

                            linhasProcessadas++;
                        }

                        // Reportar progresso do lote
                        int percentual = (linhasProcessadas * 100) / totalLinhas;
                        onProgress?.Invoke(linhasProcessadas,
                            $"Lendo Excel... {linhasProcessadas}/{totalLinhas} linhas ({percentual}%)");
                    }

                    if (linhas.Count == 0)
                    {
                        validacao.AdicionarErro("Nenhuma linha válida encontrada no arquivo");
                    }
                    else
                    {
                        onProgress?.Invoke(totalLinhas, $"Leitura concluída! {linhas.Count} linhas válidas");
                    }
                }

                return (linhas, validacao);
            }
            catch (Exception ex)
            {
                validacao.AdicionarErro($"Erro ao abrir arquivo: {ex.Message}");
                return (linhas, validacao);
            }
        }

        // Adicionar este método na classe ImportacaoController

        /// <summary>
        /// Salva as linhas importadas no banco com callback de progresso
        /// </summary>
        /// <param name="grupoCode">Código do grupo</param>
        /// <param name="linhas">Lista de linhas a serem salvas</param>
        /// <param name="onProgress">Callback para reportar progresso (linhasProcessadas, mensagem)</param>
        public void SalvarLinhasComProgresso(string grupoCode, List<LinhaImportacao> linhas, Action<int, string> onProgress)
        {
            UserTable userTable = null;
            Recordset oRecordset = null;

            try
            {
                userTable = (UserTable)_company.UserTables.Item("IT_LINHA_LOTE");

                int totalLinhas = linhas.Count;
                int linhasProcessadas = 0;
                int ultimoPercentualReportado = 0;

                // Reportar início
                onProgress?.Invoke(0, $"Iniciando salvamento de {totalLinhas} linhas...");

                foreach (var linha in linhas)
                {
                    string code = $"{grupoCode}_{linha.NumeroLinha:D4}";

                    userTable.Name = code;
                    userTable.UserFields.Fields.Item("U_GrupoCode").Value = grupoCode;
                    userTable.UserFields.Fields.Item("U_NumeroLinha").Value = linha.NumeroLinha;
                    userTable.UserFields.Fields.Item("U_Filial").Value = linha.Filial ?? "";
                    userTable.UserFields.Fields.Item("U_CodCliente").Value = linha.CodigoCliente;
                    userTable.UserFields.Fields.Item("U_NomeCliente").Value = linha.NomeCliente;
                    userTable.UserFields.Fields.Item("U_CodItem").Value = linha.CodigoItem;
                    userTable.UserFields.Fields.Item("U_DescItem").Value = linha.DescricaoItem;
                    userTable.UserFields.Fields.Item("U_Utilizacao").Value = linha.Utilizacao ?? "";
                    userTable.UserFields.Fields.Item("U_CodImposto").Value = linha.CodigoImposto ?? "";
                    userTable.UserFields.Fields.Item("U_SeqNF").Value = linha.CodSeq ?? "";
                    userTable.UserFields.Fields.Item("U_CondPagto").Value = linha.CondicaoPagamento ?? "";
                    userTable.UserFields.Fields.Item("U_Valor").Value = Convert.ToDouble(linha.Valor);
                    userTable.UserFields.Fields.Item("U_ObsNF").Value = linha.ObservacaoNF ?? "";
                    userTable.UserFields.Fields.Item("U_TipoTrib").Value = linha.TipoTributacao ?? "";
                    userTable.UserFields.Fields.Item("U_Status").Value = "P"; // Pendente
                    userTable.UserFields.Fields.Item("U_Reprocessar").Value = "N";

                    int ret = userTable.Add();
                    if (ret != 0)
                    {
                        throw new Exception($"Erro ao salvar linha {linha.NumeroLinha}: {_company.GetLastErrorDescription()}");
                    }

                    linha.Code = code;
                    linhasProcessadas++;

                    // Calcular percentual atual
                    int percentualAtual = (linhasProcessadas * 100) / totalLinhas;

                    // Reportar progresso de forma inteligente:
                    // - A cada 10 linhas para arquivos pequenos
                    // - A cada 5% para arquivos grandes
                    // - Sempre na última linha
                    bool deveReportarProgresso = false;

                    if (totalLinhas <= 100)
                    {
                        // Para arquivos pequenos, reportar a cada 10 linhas
                        deveReportarProgresso = linhasProcessadas % 10 == 0;
                    }
                    else
                    {
                        // Para arquivos grandes, reportar a cada 5%
                        deveReportarProgresso = percentualAtual >= ultimoPercentualReportado + 5;
                    }

                    // Sempre reportar a última linha
                    if (linhasProcessadas == totalLinhas)
                    {
                        deveReportarProgresso = true;
                    }

                    if (deveReportarProgresso)
                    {
                        onProgress?.Invoke(linhasProcessadas, $"Salvando... {linhasProcessadas}/{totalLinhas} linhas ({percentualAtual}%)");
                        ultimoPercentualReportado = percentualAtual;
                    }
                }

                // Atualizar o total de linhas no grupo usando SQL
                onProgress?.Invoke(linhasProcessadas, "Atualizando totalizadores...");

                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string updateQuery = $@"
            UPDATE ""@IT_GRUPO_LOTE"" 
            SET ""U_TotalLinhas"" = {linhas.Count}
            WHERE ""Code"" = '{grupoCode}'";

                oRecordset.DoQuery(updateQuery);

                // Reportar conclusão
                onProgress?.Invoke(totalLinhas, "Salvamento concluído!");
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar linhas: {ex.Message}");
            }
            finally
            {
                if (userTable != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(userTable);
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        // OPCIONAL: Versão otimizada com SQL para grandes volumes
        /// <summary>
        /// Salva as linhas usando SQL direto com callback de progresso (mais rápido para muitas linhas)
        /// </summary>
        public void SalvarLinhasSQLComProgresso(string grupoCode, List<LinhaImportacao> linhas, Action<int, string> onProgress)
        {
            Recordset oRecordset = null;

            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                int totalLinhas = linhas.Count;
                int linhasProcessadas = 0;
                int loteSize = 50; // Processar em lotes de 50 para melhor performance

                onProgress?.Invoke(0, $"Preparando para salvar {totalLinhas} linhas...");

                // Processar em lotes
                for (int i = 0; i < linhas.Count; i += loteSize)
                {
                    var lote = linhas.Skip(i).Take(loteSize).ToList();
                    var queryBuilder = new StringBuilder();

                    foreach (var linha in lote)
                    {
                        string code = $"{grupoCode}_{linha.NumeroLinha:D4}";

                        // Escapar strings para SQL
                        string nomeCliente = linha.NomeCliente.Replace("'", "''");
                        string descItem = linha.DescricaoItem.Replace("'", "''");
                        string obsNF = (linha.ObservacaoNF ?? "").Replace("'", "''");

                        queryBuilder.AppendLine($@"
                    INSERT INTO ""@IT_LINHA_LOTE"" 
                    (""Code"", ""Name"", ""U_GrupoCode"", ""U_NumeroLinha"", ""U_Filial"",
                     ""U_CodCliente"", ""U_NomeCliente"", ""U_CodItem"", ""U_DescItem"", 
                     ""U_Utilizacao"", ""U_CodImposto"", ""U_SeqNF"", ""U_CondPagto"", 
                     ""U_Valor"", ""U_ObsNF"", ""U_TipoTrib"", ""U_Status"", ""U_Reprocessar"")
                    VALUES 
                    ('{code}', '{code}', '{grupoCode}', {linha.NumeroLinha}, '{linha.Filial ?? ""}',
                     '{linha.CodigoCliente}', '{nomeCliente}', '{linha.CodigoItem}', '{descItem}',
                     '{linha.Utilizacao ?? ""}', '{linha.CodigoImposto ?? ""}', '{linha.CodSeq ?? ""}',
                     '{linha.CondicaoPagamento ?? "0"}', {linha.Valor.ToString(System.Globalization.CultureInfo.InvariantCulture)},
                     '{obsNF}', '{linha.TipoTributacao ?? ""}', 'P', 'N');");

                        linha.Code = code;
                    }

                    // Executar o lote
                    oRecordset.DoQuery(queryBuilder.ToString());

                    linhasProcessadas += lote.Count;
                    int percentual = (linhasProcessadas * 100) / totalLinhas;
                    onProgress?.Invoke(linhasProcessadas, $"Salvando... {linhasProcessadas}/{totalLinhas} linhas ({percentual}%)");
                }

                // Atualizar o total de linhas no grupo
                onProgress?.Invoke(totalLinhas, "Atualizando totalizadores...");

                string updateQuery = $@"
            UPDATE ""@IT_GRUPO_LOTE"" 
            SET ""U_TotalLinhas"" = {linhas.Count}
            WHERE ""Code"" = '{grupoCode}'";

                oRecordset.DoQuery(updateQuery);

                onProgress?.Invoke(totalLinhas, "Salvamento concluído!");
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar linhas: {ex.Message}");
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        // ALTERNATIVA: Se quiser usar SQL direto (mais rápido para muitas linhas)
        public void SalvarLinhasSQL(string grupoCode, List<LinhaImportacao> linhas)
        {
            Recordset oRecordset = null;

            try
            {
                oRecordset = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                foreach (var linha in linhas)
                {
                    string code = $"{grupoCode}_{linha.NumeroLinha:D4}";

                    // Escapar strings para SQL
                    string nomeCliente = linha.NomeCliente.Replace("'", "''");
                    string descItem = linha.DescricaoItem.Replace("'", "''");

                    string insertQuery = $@"
                INSERT INTO ""@IT_LINHA_LOTE"" 
                (""Code"", ""Name"", ""U_GrupoCode"", ""U_NumeroLinha"", ""U_CodCliente"", 
                 ""U_NomeCliente"", ""U_CodItem"", ""U_DescItem"", ""U_Utilizacao"", 
                 ""U_CodImposto"", ""U_SeqNF"", ""U_CondPagto"", ""U_Valor"", 
                 ""U_Status"", ""U_Reprocessar"")
                VALUES 
                ('{code}', '{code}', '{grupoCode}', {linha.NumeroLinha}, '{linha.CodigoCliente}',
                 '{nomeCliente}', '{linha.CodigoItem}', '{descItem}', '{linha.Utilizacao ?? ""}',
                 '{linha.CodigoImposto ?? ""}', '{linha.CodSeq ?? ""}', 
                 '{linha.CondicaoPagamento ?? "0"}', {linha.Valor.ToString(System.Globalization.CultureInfo.InvariantCulture)},
                 'P', 'N')";

                    oRecordset.DoQuery(insertQuery);

                    linha.Code = code;
                }

                // Atualizar o total de linhas no grupo
                string updateQuery = $@"
            UPDATE ""@IT_GRUPO_LOTE"" 
            SET ""U_TotalLinhas"" = {linhas.Count}
            WHERE ""Code"" = '{grupoCode}'";

                oRecordset.DoQuery(updateQuery);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar linhas: {ex.Message}");
            }
            finally
            {
                if (oRecordset != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            }
        }

        #region Validações

        private void ValidarLinha(LinhaImportacao linha, int numeroLinha, ValidacaoImportacao validacao)
        {
            // Validar campos obrigatórios
            if (string.IsNullOrEmpty(linha.CodigoCliente))
                validacao.AdicionarErro($"Linha {numeroLinha}: Código do cliente não informado");

            if (string.IsNullOrEmpty(linha.CodigoItem))
                validacao.AdicionarErro($"Linha {numeroLinha}: Código do item não informado");

            if (linha.Valor <= 0)
                validacao.AdicionarErro($"Linha {numeroLinha}: Valor deve ser maior que zero");

            // Validar se cliente existe
            if (!string.IsNullOrEmpty(linha.CodigoCliente) && !ClienteExiste(linha.CodigoCliente))
                validacao.AdicionarErro($"Linha {numeroLinha}: Cliente {linha.CodigoCliente} não encontrado");

            // Validar se item existe
            if (!string.IsNullOrEmpty(linha.CodigoItem) && !ItemExiste(linha.CodigoItem))
                validacao.AdicionarErro($"Linha {numeroLinha}: Item {linha.CodigoItem} não encontrado");
        }

        private bool ClienteExiste(string cardCode)
        {
            BusinessPartners oBP = null;
            try
            {
                oBP = (BusinessPartners)_company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                return oBP.GetByKey(cardCode);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
            }
        }

        private bool ItemExiste(string itemCode)
        {
            Items oItem = null;
            try
            {
                oItem = (Items)_company.GetBusinessObject(BoObjectTypes.oItems);
                return oItem.GetByKey(itemCode);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
            }
        }

        #endregion
    }
}
