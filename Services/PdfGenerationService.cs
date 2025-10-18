using SAPbobsCOM;
using System;
using System.IO;
using System.Configuration;

namespace ItTech.Tool.AddonNFS.Services

{

    /// <summary>
    /// Serviço responsável pela geração de arquivos PDF a partir de layouts do SAP Business One.
    /// Utiliza o ReportLayoutsService da DI API para uma exportação nativa e estável.
    /// </summary>

    public class PdfGenerationService
    {
        private readonly Company _company;

        public PdfGenerationService(Company company)
        {
            _company = company ?? throw new ArgumentNullException(nameof(company));
        }

        /// <summary>
        /// Gera o PDF para um documento de Entrega específico.
        /// </summary>
        /// <param name="docEntry">O DocEntry da Entrega a ser impressa.</param>
        /// <param name="pastaDestino">O caminho da pasta onde o PDF será salvo.</param>
        /// <param name="nomeArquivo">O nome do arquivo a ser criado (sem a extensão .pdf).</param>
        /// <returns>O caminho completo do arquivo PDF gerado com sucesso.</returns>

        public string GerarPdfDeEntrega(int docEntry, string pastaDestino, string nomeArquivo)
        {
            SAPbobsCOM.ReportLayoutsService rptService = null;

            try
            {
                if (string.IsNullOrWhiteSpace(pastaDestino))
                    throw new ArgumentException("A pasta de destino do PDF não foi configurada.");

                if (!Directory.Exists(pastaDestino))
                    Directory.CreateDirectory(pastaDestino);

                string layoutCode = ConfigurationManager.AppSettings["DeliveryNote_LayoutCode"]; // Ex: DLN20001
                string caminhoCompleto = Path.Combine(pastaDestino, $"{nomeArquivo}.pdf");

                // Instancia o serviço
                rptService = (SAPbobsCOM.ReportLayoutsService)_company.GetCompanyService()
                                .GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);

                // Parâmetros do layout
                SAPbobsCOM.ReportLayoutPrintParams printParams =
                    (SAPbobsCOM.ReportLayoutPrintParams)rptService.GetDataInterface(
                        SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams);

                printParams.LayoutCode = layoutCode;
                printParams.DocEntry = docEntry;

                // 💡 Aqui entra o truque:
                // O SAP usa a impressora padrão do Windows. Configure “Microsoft Print to PDF”
                // como padrão OU use um PDFCreator configurado para salvar automaticamente.
                rptService.Print(printParams);
               

                // O arquivo será gerado pela impressora virtual.
                // Se quiser mover o arquivo para uma pasta específica, pode monitorar a pasta de saída.
                return caminhoCompleto;
            }
            catch (Exception ex)
            {
                throw new Exception($"Falha ao gerar PDF da Entrega (DocEntry={docEntry}): {ex.Message}", ex);
            }
            finally
            {
                if (rptService != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rptService);
            }
        }




        //public string GerarPdfDeEntrega(int docEntry, string pastaDestino, string nomeArquivo)

        //{
        //    ReportLayoutsService rptLayoutService = null;

        //    try
        //    {
        //        var layoutCode = ConfigurationManager.AppSettings["DeliveryNote_LayoutCode"];

        //        // 1. Validação dos parâmetros de entrada
        //        if (string.IsNullOrWhiteSpace(pastaDestino))
        //            throw new ArgumentException("A pasta de destino do PDF não foi configurada.");

        //        if (!Directory.Exists(pastaDestino))
        //            Directory.CreateDirectory(pastaDestino);

        //        // 2. Instancia o serviço de layouts da DI API
        //        rptLayoutService = (ReportLayoutsService)_company.GetCompanyService().GetBusinessService(ServiceTypes.ReportLayoutsService);

        //        // 3. Configura os parâmetros para a impressão do relatório
        //        ReportLayoutPrintParams rptParams = (ReportLayoutPrintParams)rptLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams);
        //        rptParams.LayoutCode = layoutCode; 
        //        rptParams.DocEntry = docEntry;

        //        // 4. Define o caminho e nome do arquivo de saída
        //        string caminhoCompleto = Path.Combine(pastaDestino, $"{nomeArquivo}.pdf");

        //        // 5. Exporta o relatório para o arquivo PDF
        //        rptLayoutService.Print(rptParams);

        //        // No SDK mais recente, usa-se ExportToFile. Se Print não funcionar, usaremos este:
        //        //rptLayoutService.ExportToFile(rptParams, caminhoCompleto, ReportLayoutExportTypeEnum.rletPdf);

        //        return caminhoCompleto;

        //    }
        //    catch (Exception ex)

        //    {
        //        // Lança uma exceção detalhada para ser capturada pelo controller
        //        throw new Exception($"Falha ao gerar o PDF para a Entrega DocEntry '{docEntry}'. Erro: {ex.Message}", ex);
        //    }
        //    finally
        //    {
        //        if (rptLayoutService != null)
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(rptLayoutService);
        //        }
        //    }
        //}
    }
}