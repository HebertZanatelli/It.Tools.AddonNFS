using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace ItTech.Tool.AddonNFS
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {

                // Supressão GLOBAL e DRASTICA de exceções no Add-on
                AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
                {
                    // Aqui você pode logar os erros suprimidos, se desejar
                    System.Diagnostics.Debug.WriteLine($"UnhandledException drasticamente suprimido: {(e.ExceptionObject as Exception)?.Message}");
                    // Não lança novamente, suprime totalmente
                };

                //Application.ThreadException += (sender, e) =>
                //{
                //    System.Diagnostics.Debug.WriteLine($"ThreadException drasticamente suprimido: {e.Exception.Message}");
                //    // Suprime drasticamente
                //};

                AppDomain.CurrentDomain.FirstChanceException += (sender, e) =>
                {
                    System.Diagnostics.Debug.WriteLine($"FirstChanceException drasticamente suprimido: {e.Exception.Message}");
                    // Suprime drasticamente
                };

                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();

                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);


                // ========================================
                // AQUI É ONDE AS TABELAS SÃO CRIADAS!
                // ========================================
                try
                {
                    Application.SBO_Application.StatusBar.SetText("Verificando estrutura do banco de dados...",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    // Criar instância do DatabaseSetup
                    var dbSetup = new Database.DatabaseSetup((Company) Application.SBO_Application.Company.GetDICompany());

                    // EXECUTAR A CRIAÇÃO DAS TABELAS E CAMPOS
                    dbSetup.CreateDatabaseStructure();

                    Application.SBO_Application.StatusBar.SetText("Add-on NFS-e em Lote iniciado com sucesso!",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                catch (Exception dbEx)
                {
                    // Se houver erro na criação das tabelas, mostrar mensagem mas continuar
                    Application.SBO_Application.MessageBox(
                        $"Aviso: Erro ao verificar estrutura do banco.\n{dbEx.Message}\n\n" +
                        "O add-on continuará funcionando, mas pode apresentar erros.",
                        1, "Ok", "", "");
                }

                //-------------------------------------------------------

                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
