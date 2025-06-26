using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ItTech.Tool.AddonNFS.Models;

namespace ItTech.Tool.AddonNFS.Utils
{
    /// <summary>
    /// Serviço para exportação de erros de processamento NFS-e para Excel
    /// </summary>
    public class ExcelExportService
    {
        private readonly string _exportPath;

        public ExcelExportService()
        {
            // Define pasta de exportação (Documentos do usuário)
            _exportPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "NFS-e Exportações"
            );

            // Criar pasta se não existir
            if (!Directory.Exists(_exportPath))
            {
                Directory.CreateDirectory(_exportPath);
            }
        }

        /// <summary>
        /// Exporta linhas com erro mantendo o layout original + coluna de mensagem de erro
        /// </summary>
        /// <param name="grupo">Grupo de lote processado</param>
        /// <param name="linhasComErro">Lista de linhas com erro</param>
        /// <returns>Caminho completo do arquivo gerado</returns>
        public string ExportarErros(GrupoLote grupo, List<LinhaImportacao> linhasComErro, string caminhoCompleto)
        {
            try
            {
                // Criar workbook
                using (var workbook = new XLWorkbook())
                {
                    // Adicionar planilha
                    var worksheet = workbook.Worksheets.Add("Erros NFS-e");

                    // Configurar cabeçalhos (mantendo layout original)
                    ConfigurarCabecalhos(worksheet);

                    // Adicionar dados das linhas com erro
                    int linha = 2; // Começa na linha 2 (após cabeçalho)
                    foreach (var linhaErro in linhasComErro)
                    {
                        PreencherLinha(worksheet, linha, linhaErro);
                        linha++;
                    }

                    // Formatar planilha
                    FormatarPlanilha(worksheet, linhasComErro.Count);

                    // Gerar nome do arquivo
                    string nomeArquivo = $"NFS-e_Erros_{grupo.Nome}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    //string caminhoCompleto = Path.Combine(_exportPath, nomeArquivo);

                    // Salvar arquivo
                    workbook.SaveAs(caminhoCompleto);

                    return caminhoCompleto;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao exportar erros: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Exporta template vazio para nova importação
        /// </summary>
        public string ExportarTemplate()
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Template NFS-e");

                    // Apenas cabeçalhos
                    ConfigurarCabecalhos(worksheet, incluirMensagemErro: false);

                    // Adicionar linha de exemplo
                    worksheet.Cell("A2").Value = 1;
                    worksheet.Cell("B2").Value = "C00001";
                    worksheet.Cell("C2").Value = "Cliente Exemplo LTDA";
                    worksheet.Cell("D2").Value = "S00001";
                    worksheet.Cell("E2").Value = "Serviço de Exemplo";
                    worksheet.Cell("F2").Value = 13;
                    worksheet.Cell("G2").Value = "1124-001";
                    worksheet.Cell("H2").Value = 31;
                    worksheet.Cell("I2").Value = -2;
                    worksheet.Cell("J2").Value = 1000.00;
                    worksheet.Cell("K2").Value = "Observação exemplo";
                    worksheet.Cell("L2").Value = 1;

                    // Formatar
                    FormatarPlanilha(worksheet, 1, incluirMensagemErro: false);

                    // Adicionar comentários explicativos
                    AdicionarComentariosTemplate(worksheet);

                    string nomeArquivo = $"Template_NFS-e_{DateTime.Now:yyyyMMdd}.xlsx";
                    string caminhoCompleto = Path.Combine(_exportPath, nomeArquivo);

                    workbook.SaveAs(caminhoCompleto);
                    return caminhoCompleto;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao exportar template: {ex.Message}", ex);
            }
        }

        private void ConfigurarCabecalhos(IXLWorksheet worksheet, bool incluirMensagemErro = true)
        {
            // Cabeçalhos originais (A-L)
            worksheet.Cell("A1").Value = "Filial";
            worksheet.Cell("B1").Value = "Código do Cliente";
            worksheet.Cell("C1").Value = "Nome do Cliente";
            worksheet.Cell("D1").Value = "Código do Item";
            worksheet.Cell("E1").Value = "Descrição do Item";
            worksheet.Cell("F1").Value = "Utilização";
            worksheet.Cell("G1").Value = "Codigo Imposto";
            worksheet.Cell("H1").Value = "Cod Seq";
            worksheet.Cell("I1").Value = "Condição de Pagamento";
            worksheet.Cell("J1").Value = "Valor";
            worksheet.Cell("K1").Value = "OBS NF";
            worksheet.Cell("L1").Value = "Tipo Tributação";

            // Coluna adicional para mensagem de erro
            if (incluirMensagemErro)
            {
                worksheet.Cell("M1").Value = "❌ MENSAGEM DE ERRO";
                worksheet.Cell("M1").Style.Fill.BackgroundColor = XLColor.Red;
                worksheet.Cell("M1").Style.Font.FontColor = XLColor.White;
                worksheet.Cell("M1").Style.Font.Bold = true;
            }

            // Estilo dos cabeçalhos
            var rangeCabecalho = incluirMensagemErro
                ? worksheet.Range("A1:M1")
                : worksheet.Range("A1:L1");

            rangeCabecalho.Style.Font.Bold = true;
            rangeCabecalho.Style.Fill.BackgroundColor = XLColor.FromHtml("#003366");
            rangeCabecalho.Style.Font.FontColor = XLColor.White;
            rangeCabecalho.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangeCabecalho.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        private void PreencherLinha(IXLWorksheet worksheet, int linha, LinhaImportacao linhaErro)
        {
            // Dados originais (A-L)
            worksheet.Cell($"A{linha}").Value = linhaErro.Filial;
            worksheet.Cell($"B{linha}").Value = linhaErro.CodigoCliente;
            worksheet.Cell($"C{linha}").Value = linhaErro.NomeCliente;
            worksheet.Cell($"D{linha}").Value = linhaErro.CodigoItem;
            worksheet.Cell($"E{linha}").Value = linhaErro.DescricaoItem;
            worksheet.Cell($"F{linha}").Value = linhaErro.Utilizacao;
            worksheet.Cell($"G{linha}").Value = linhaErro.CodigoImposto;
            worksheet.Cell($"H{linha}").Value = linhaErro.CodSeq;
            worksheet.Cell($"I{linha}").Value = linhaErro.CondicaoPagamento;
            worksheet.Cell($"J{linha}").Value = linhaErro.Valor;
            worksheet.Cell($"K{linha}").Value = linhaErro.ObservacaoNF;
            worksheet.Cell($"L{linha}").Value = linhaErro.TipoTributacao;

            // Mensagem de erro (M)
            worksheet.Cell($"M{linha}").Value = linhaErro.MensagemErro;
            worksheet.Cell($"M{linha}").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFE6E6");
            worksheet.Cell($"M{linha}").Style.Font.FontColor = XLColor.DarkRed;
            worksheet.Cell($"M{linha}").Style.Alignment.WrapText = true;

            // Destacar linha com erro
            var rangeErro = worksheet.Range($"A{linha}:L{linha}");
            rangeErro.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF0F0");
        }

        private void FormatarPlanilha(IXLWorksheet worksheet, int totalLinhas, bool incluirMensagemErro = true)
        {
            // Auto ajustar largura das colunas
            worksheet.Columns().AdjustToContents();

            // Larguras específicas
            worksheet.Column("A").Width = 8;   // Filial
            worksheet.Column("B").Width = 15;  // Código Cliente
            worksheet.Column("C").Width = 35;  // Nome Cliente
            worksheet.Column("D").Width = 15;  // Código Item
            worksheet.Column("E").Width = 35;  // Descrição Item
            worksheet.Column("F").Width = 12;  // Utilização
            worksheet.Column("G").Width = 15;  // Código Imposto
            worksheet.Column("H").Width = 10;  // Cod Seq
            worksheet.Column("I").Width = 20;  // Condição Pagamento
            worksheet.Column("J").Width = 15;  // Valor
            worksheet.Column("K").Width = 30;  // OBS NF
            worksheet.Column("L").Width = 15;  // Tipo Tributação

            if (incluirMensagemErro)
            {
                worksheet.Column("M").Width = 50; // Mensagem Erro
            }

            // Formatar coluna de valor como moeda
            var rangeValor = worksheet.Range($"J2:J{totalLinhas + 1}");
            rangeValor.Style.NumberFormat.Format = "R$ #,##0.00";

            // Aplicar bordas
            var rangeTotal = incluirMensagemErro
                ? worksheet.Range($"A1:M{totalLinhas + 1}")
                : worksheet.Range($"A1:L{totalLinhas + 1}");

            rangeTotal.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            rangeTotal.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Congelar painel (primeira linha)
            worksheet.SheetView.FreezeRows(1);

            // Adicionar filtros
            rangeTotal.SetAutoFilter();

           
        }

        private void AdicionarComentariosTemplate(IXLWorksheet worksheet)
        {
            // Adicionar comentários explicativos em cada célula do cabeçalho
            // Sintaxe correta para ClosedXML
            worksheet.Cell("A1").Comment.AddText("Código da filial (número inteiro)");
            worksheet.Cell("B1").Comment.AddText("Código do cliente no SAP B1 (ex: C00001)");
            worksheet.Cell("C1").Comment.AddText("Nome completo do cliente");
            worksheet.Cell("D1").Comment.AddText("Código do item de serviço no SAP B1 (ex: S00001)");
            worksheet.Cell("E1").Comment.AddText("Descrição completa do serviço");
            worksheet.Cell("F1").Comment.AddText("Código de utilização (geralmente 13 para serviços)");
            worksheet.Cell("G1").Comment.AddText("Código do imposto no SAP B1 (ex: 1124-001)");
            worksheet.Cell("H1").Comment.AddText("Código da sequência de numeração NF");
            worksheet.Cell("I1").Comment.AddText("Código da condição de pagamento (-2 = à vista, ou código específico)");
            worksheet.Cell("J1").Comment.AddText("Valor do serviço (formato: 9999.99)");
            worksheet.Cell("K1").Comment.AddText("Observações que aparecerão na NF (máx 254 caracteres)");
            worksheet.Cell("L1").Comment.AddText("Tipo de tributação (1, 4, 5, F, etc.)");

            // Configurar tamanho dos comentários
            // Em versões mais recentes do ClosedXML, o tamanho é ajustado automaticamente
            // Se precisar ajustar manualmente, use:
            worksheet.Cell("A1").Comment.Style.Size.SetHeight(40);
            worksheet.Cell("A1").Comment.Style.Size.SetWidth(200);
                
            // Aplicar o mesmo estilo para todos os comentários
            for (int col = 2; col <= 12; col++) // B até L
            {
                var cell = worksheet.Cell(1, col);
                if (cell.HasComment)
                {
                    cell.Comment.Style.Size.SetHeight(40);
                    cell.Comment.Style.Size.SetWidth(200);
                }
            }
        }

        /// <summary>
        /// Exporta relatório consolidado do processamento
        /// </summary>
        public string ExportarRelatorioProcessamento(GrupoLote grupo)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    // Planilha 1: Resumo
                    var wsResumo = workbook.Worksheets.Add("Resumo");
                    CriarPlanilhaResumo(wsResumo, grupo);

                    // Planilha 2: Sucessos
                    var wsSucessos = workbook.Worksheets.Add("Sucessos");
                    CriarPlanilhaDetalhes(wsSucessos, grupo.Linhas.Where(l => l.Status == StatusLinha.Sucesso).ToList(), "Sucesso");

                    // Planilha 3: Erros
                    var wsErros = workbook.Worksheets.Add("Erros");
                    CriarPlanilhaDetalhes(wsErros, grupo.Linhas.Where(l => l.Status == StatusLinha.Erro).ToList(), "Erro");

                    // Planilha 4: Pendentes
                    var wsPendentes = workbook.Worksheets.Add("Pendentes");
                    CriarPlanilhaDetalhes(wsPendentes, grupo.Linhas.Where(l => l.Status == StatusLinha.Pendente).ToList(), "Pendente");

                    string nomeArquivo = $"Relatorio_NFS-e_{grupo.Nome}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    string caminhoCompleto = Path.Combine(_exportPath, nomeArquivo);

                    workbook.SaveAs(caminhoCompleto);
                    return caminhoCompleto;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao exportar relatório: {ex.Message}", ex);
            }
        }

        private void CriarPlanilhaResumo(IXLWorksheet worksheet, GrupoLote grupo)
        {
            // Título
            worksheet.Cell("A1").Value = "RELATÓRIO DE PROCESSAMENTO NFS-e";
            worksheet.Cell("A1").Style.Font.Bold = true;
            worksheet.Cell("A1").Style.Font.FontSize = 16;
            worksheet.Range("A1:D1").Merge();

            // Informações do grupo
            worksheet.Cell("A3").Value = "Grupo:";
            worksheet.Cell("B3").Value = grupo.Nome;
            worksheet.Cell("A4").Value = "Data Lançamento:";
            worksheet.Cell("B4").Value = grupo.DataLancamento.ToString("dd/MM/yyyy");
            worksheet.Cell("A5").Value = "Data Documento:";
            worksheet.Cell("B5").Value = grupo.DataDocumento.ToString("dd/MM/yyyy");

            // Estatísticas
            int totalLinhas = grupo.Linhas.Count;
            int sucessos = grupo.Linhas.Count(l => l.Status == StatusLinha.Sucesso);
            int erros = grupo.Linhas.Count(l => l.Status == StatusLinha.Erro);
            int pendentes = grupo.Linhas.Count(l => l.Status == StatusLinha.Pendente);
            decimal valorTotal = grupo.Linhas.Sum(l => l.Valor);
            decimal valorSucesso = grupo.Linhas.Where(l => l.Status == StatusLinha.Sucesso).Sum(l => l.Valor);

            worksheet.Cell("A7").Value = "ESTATÍSTICAS";
            worksheet.Cell("A7").Style.Font.Bold = true;

            worksheet.Cell("A8").Value = "Total de Documentos:";
            worksheet.Cell("B8").Value = totalLinhas;

            worksheet.Cell("A9").Value = "✅ Sucessos:";
            worksheet.Cell("B9").Value = sucessos;
            worksheet.Cell("C9").Value = $"{(sucessos * 100.0 / totalLinhas):F1}%";

            worksheet.Cell("A10").Value = "❌ Erros:";
            worksheet.Cell("B10").Value = erros;
            worksheet.Cell("C10").Value = $"{(erros * 100.0 / totalLinhas):F1}%";

            worksheet.Cell("A11").Value = "⏳ Pendentes:";
            worksheet.Cell("B11").Value = pendentes;
            worksheet.Cell("C11").Value = $"{(pendentes * 100.0 / totalLinhas):F1}%";

            worksheet.Cell("A13").Value = "Valor Total:";
            worksheet.Cell("B13").Value = valorTotal;
            worksheet.Cell("B13").Style.NumberFormat.Format = "R$ #,##0.00";

            worksheet.Cell("A14").Value = "Valor Processado:";
            worksheet.Cell("B14").Value = valorSucesso;
            worksheet.Cell("B14").Style.NumberFormat.Format = "R$ #,##0.00";

            // Formatar
            worksheet.Columns().AdjustToContents();
        }

        private void CriarPlanilhaDetalhes(IXLWorksheet worksheet, List<LinhaImportacao> linhas, string tipo)
        {
            // Cabeçalhos básicos + DocEntry, DocNum, NumNF
            worksheet.Cell("A1").Value = "Linha";
            worksheet.Cell("B1").Value = "Cliente";
            worksheet.Cell("C1").Value = "Item";
            worksheet.Cell("D1").Value = "Valor";
            worksheet.Cell("E1").Value = "DocEntry";
            worksheet.Cell("F1").Value = "DocNum";
            worksheet.Cell("G1").Value = "Número NF";

            if (tipo == "Erro")
            {
                worksheet.Cell("H1").Value = "Mensagem de Erro";
            }

            // Estilo cabeçalho
            var rangeCabecalho = tipo == "Erro" ? worksheet.Range("A1:H1") : worksheet.Range("A1:G1");
            rangeCabecalho.Style.Font.Bold = true;
            rangeCabecalho.Style.Fill.BackgroundColor = XLColor.FromHtml("#003366");
            rangeCabecalho.Style.Font.FontColor = XLColor.White;

            // Dados
            int linha = 2;
            foreach (var item in linhas)
            {
                worksheet.Cell($"A{linha}").Value = item.NumeroLinha;
                worksheet.Cell($"B{linha}").Value = $"{item.CodigoCliente} - {item.NomeCliente}";
                worksheet.Cell($"C{linha}").Value = $"{item.CodigoItem} - {item.DescricaoItem}";
                worksheet.Cell($"D{linha}").Value = item.Valor;
                worksheet.Cell($"E{linha}").Value = item.DocEntry ?? 0;
                worksheet.Cell($"F{linha}").Value = item.DocNum ?? 0;

                if (tipo == "Erro")
                {
                    worksheet.Cell($"H{linha}").Value = item.MensagemErro;
                }

                linha++;
            }

            // Formatar valores
            var rangeValor = worksheet.Range($"D2:D{linha - 1}");
            rangeValor.Style.NumberFormat.Format = "R$ #,##0.00";

            // Total
            worksheet.Cell($"C{linha + 1}").Value = "TOTAL:";
            worksheet.Cell($"C{linha + 1}").Style.Font.Bold = true;
            worksheet.Cell($"D{linha + 1}").FormulaA1 = $"=SUM(D2:D{linha - 1})";
            worksheet.Cell($"D{linha + 1}").Style.NumberFormat.Format = "R$ #,##0.00";
            worksheet.Cell($"D{linha + 1}").Style.Font.Bold = true;

            // Ajustar colunas
            worksheet.Columns().AdjustToContents();
            worksheet.Column("B").Width = 40;
            worksheet.Column("C").Width = 40;

            if (tipo == "Erro")
            {
                worksheet.Column("H").Width = 50;
            }

            // Congelar primeira linha
            worksheet.SheetView.FreezeRows(1);
        }
    }
}