using SAPbouiCOM.Framework;
using System;
using ItTech.Tool.AddonNFS.Forms;
using ItTech.Tool.AddonNFS.Utils;

namespace ItTech.Tool.AddonNFS
{
    /// <summary>
    /// Classe responsável pelo gerenciamento dos menus do add-on
    /// </summary>
    class Menu
    {
        /// <summary>
        /// Adiciona os itens de menu do add-on ao SAP Business One
        /// </summary>
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            try
            {
                oMenus = Application.SBO_Application.Menus;

                SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                // Obter menu de módulos
                oMenuItem = Application.SBO_Application.Menus.Item("43520"); // modules

                // Configurar menu principal
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "ITTECH_NFS";
                oCreationPackage.String = "NFS-e em Lote";
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = -1;

                oMenus = oMenuItem.SubMenus;

                try
                {
                    // Se o menu já existe, remove primeiro
                    if (Application.SBO_Application.Menus.Exists("ITTECH_NFS"))
                    {
                        Application.SBO_Application.Menus.RemoveEx("ITTECH_NFS");
                    }

                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception e)
                {
                    // Menu já existe, continuar
                    Application.SBO_Application.SetStatusBarMessage($"Menu principal já existe: {e.Message}",
                        SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }

                // Adicionar submenus
                AdicionarSubmenus(oCreationPackage);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao criar menus: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Adiciona os submenus ao menu principal
        /// </summary>
        private void AdicionarSubmenus(SAPbouiCOM.MenuCreationParams oCreationPackage)
        {
            try
            {
                // Obter o menu collection do item popup recém adicionado
                var oMenuItem = Application.SBO_Application.Menus.Item("ITTECH_NFS");
                var oMenus = oMenuItem.SubMenus;

                // Criar submenu - Processar NFS-e em Lote
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ITTECH_NFS_PROC";
                oCreationPackage.String = "Processar NFS-e em Lote";
                if (!Application.SBO_Application.Menus.Exists("ITTECH_NFS_PROC"))
                {
                    oMenus.AddEx(oCreationPackage);
                }

                //// Criar submenu - Configurações
                //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //oCreationPackage.UniqueID = "ITTECH_NFS_CONFIG";
                //oCreationPackage.String = "Configurações";
                //if (!Application.SBO_Application.Menus.Exists("ITTECH_NFS_CONFIG"))
                //{
                //    oMenus.AddEx(oCreationPackage);
                //}

                //// Criar submenu - Status
                //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //oCreationPackage.UniqueID = "ITTECH_NFS_STATUS";
                //oCreationPackage.String = "Status dos Formulários";
                //if (!Application.SBO_Application.Menus.Exists("ITTECH_NFS_STATUS"))
                //{
                //    oMenus.AddEx(oCreationPackage);
                //}

                //// Criar submenu - Sobre
                //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //oCreationPackage.UniqueID = "ITTECH_NFS_ABOUT";
                //oCreationPackage.String = "Sobre";
                if (!Application.SBO_Application.Menus.Exists("ITTECH_NFS_ABOUT"))
                {
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception er)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao criar submenus: {er.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Manipula os eventos de menu
        /// </summary>
        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "ITTECH_NFS_PROC":
                            AbrirFormularioSelecionarGrupo();
                            break;

                        //case "ITTECH_NFS_CONFIG":
                        //    AbrirFormularioConfiguracao();
                        //    break;

                        //case "ITTECH_NFS_STATUS":
                        //    MostrarStatusFormularios();
                        //    break;

                        //case "ITTECH_NFS_ABOUT":
                        //    MostrarSobre();
                        //    break;
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro no evento de menu: {ex.Message}", 1, "Ok", "", "");
            }
        }

        /// <summary>
        /// Abre o formulário de seleção de grupo (Formulário Mestre)
        /// </summary>
        private void AbrirFormularioSelecionarGrupo()
        {
            try
            {
                // Verificar se já está aberto usando FormManager
                if (FormManager.FormularioEstaAberto(FormManager.FORM_SELECIONAR_GRUPO))
                {
                    // Trazer para frente
                    FormManager.TrazerParaFrente(FormManager.FORM_SELECIONAR_GRUPO);

                    Application.SBO_Application.SetStatusBarMessage("Formulário de seleção já está aberto",
                        SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    // Verificar se há algum formulário filho aberto
                    if (FormManager.ExistemFormulariosFilhosAbertos())
                    {
                        int resposta = Application.SBO_Application.MessageBox(
                            "Existem formulários de processamento abertos. Deseja fechá-los?",
                            2, "Sim", "Não", "");

                        if (resposta == 1)
                        {
                            // Fechar todos os formulários
                            FormManager.FecharTodosFormularios();
                        }
                    }

                    // Criar nova instância
                    FormSelecionarGrupo formSelecao = new FormSelecionarGrupo();
                    formSelecao.Show();

                    // O registro no FormManager deve ser feito dentro do próprio formulário
                    // após a inicialização completa
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao abrir formulário: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Abre o formulário de configurações
        /// </summary>
        private void AbrirFormularioConfiguracao()
        {
            try
            {
                if (FormManager.FormularioEstaAberto(FormManager.FORM_CONFIG))
                {
                    FormManager.TrazerParaFrente(FormManager.FORM_CONFIG);
                    Application.SBO_Application.SetStatusBarMessage("Formulário de configurações já está aberto",
                        SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    // Por enquanto, apenas mostrar mensagem
                    Application.SBO_Application.MessageBox(
                        "Configurações em desenvolvimento\n\n" +
                        "Funcionalidades previstas:\n" +
                        "- Configuração de caminhos padrão\n" +
                        "- Parâmetros de processamento\n" +
                        "- Configurações de Service Layer\n" +
                        "- Logs e diagnósticos",
                        1, "Ok", "", "");

                    // TODO: Quando implementar:
                    // FormConfiguracoes formConfig = new FormConfiguracoes();
                    // formConfig.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Mostra o status dos formulários abertos
        /// </summary>
        private void MostrarStatusFormularios()
        {
            try
            {
                string status = FormManager.ObterStatusFormularios();

                if (string.IsNullOrEmpty(status) || status.Contains("Formulários abertos: 0"))
                {
                    Application.SBO_Application.MessageBox(
                        "Nenhum formulário do add-on está aberto no momento.\n\n" +
                        "Use o menu 'Processar NFS-e em Lote' para iniciar.",
                        1, "Ok", "", "");
                }
                else
                {
                    Application.SBO_Application.MessageBox(status, 1, "Ok", "", "");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao obter status: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Mostra informações sobre o add-on
        /// </summary>
        private void MostrarSobre()
        {
            try
            {
                // Obter um resumo simples dos formulários
                string statusFormularios = FormManager.ObterStatusFormularios();
                string resumoFormularios = "";

                if (!string.IsNullOrEmpty(statusFormularios))
                {
                    // Extrair apenas a primeira linha com contagem
                    var linhas = statusFormularios.Split('\n');
                    if (linhas.Length > 0)
                    {
                        resumoFormularios = "\n\n" + linhas[0];
                    }
                }

                string sobre = "Add-on NFS-e em Lote\n" +
                             "━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                             "Versão: 1.0.0\n" +
                             "Build: 2025.01.001\n\n" +
                             "Desenvolvido por: ItTech Consultoria\n" +
                             "Cliente: Tools\n\n" +
                             "Descrição:\n" +
                             "Este add-on permite a emissão de Notas Fiscais de Serviço em lote " +
                             "através da importação de planilhas Excel. Processa até 500 documentos " +
                             "com recursos de reprocessamento e controle de erros.\n\n" +
                             "Funcionalidades:\n" +
                             "• Importação de planilhas Excel\n" +
                             "• Processamento em lote via Service Layer\n" +
                             "• Reprocessamento de linhas com erro\n" +
                             "• Histórico de processamentos" +
                             resumoFormularios;

                Application.SBO_Application.MessageBox(sobre, 1, "Ok", "", "");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao mostrar sobre: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Remove os menus do add-on (para usar na desinstalação)
        /// </summary>
        public void RemoveMenuItems()
        {
            try
            {
                // Fechar todos os formulários abertos
                FormManager.FecharTodosFormularios();

                // Remover submenus
                string[] submenus = { "ITTECH_NFS_PROC", "ITTECH_NFS_CONFIG", "ITTECH_NFS_STATUS", "ITTECH_NFS_ABOUT" };

                foreach (string menuId in submenus)
                {
                    if (Application.SBO_Application.Menus.Exists(menuId))
                    {
                        Application.SBO_Application.Menus.RemoveEx(menuId);
                    }
                }

                // Remover menu principal
                if (Application.SBO_Application.Menus.Exists("ITTECH_NFS"))
                {
                    Application.SBO_Application.Menus.RemoveEx("ITTECH_NFS");
                }

                Application.SBO_Application.SetStatusBarMessage("Menus do add-on removidos com sucesso",
                    SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao remover menus: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
    }
}