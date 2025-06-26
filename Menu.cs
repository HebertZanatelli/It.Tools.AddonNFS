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
                    // Verificar se o menu não existe antes de adicionar
                    if (!Application.SBO_Application.Menus.Exists("ITTECH_NFS"))
                    {
                        oMenus.AddEx(oCreationPackage);
                    }
                    else
                    {
                        // Menu já existe, apenas continuar sem erro
                        System.Diagnostics.Debug.WriteLine("Menu ITTECH_NFS já existe, continuando...");
                    }
                }
                catch (Exception e)
                {
                    // Log silencioso para não aparecer na instalação
                    System.Diagnostics.Debug.WriteLine($"Aviso ao criar menu: {e.Message}");
                }

                // Adicionar submenus
                AdicionarSubmenus(oCreationPackage);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao criar menus: {ex.Message}");
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
            }
            catch (Exception er)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao criar submenus: {er.Message}");
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
        /// Remove os menus do add-on (para usar na desinstalação)
        /// </summary>
        public void RemoveMenuItems()
        {
            try
            {
                // Fechar todos os formulários abertos
                FormManager.FecharTodosFormularios();

                // Remover submenus
                if (Application.SBO_Application.Menus.Exists("ITTECH_NFS_PROC"))
                {
                    Application.SBO_Application.Menus.RemoveEx("ITTECH_NFS_PROC");
                }

                // Remover menu principal
                if (Application.SBO_Application.Menus.Exists("ITTECH_NFS"))
                {
                    Application.SBO_Application.Menus.RemoveEx("ITTECH_NFS");
                }

                System.Diagnostics.Debug.WriteLine("Menus do add-on removidos");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao remover menus: {ex.Message}");
            }
        }
    }
}