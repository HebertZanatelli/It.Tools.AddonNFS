using System;
using System.Collections.Generic;
using System.Linq;
using SAPbouiCOM.Framework;
using ItTech.Tool.AddonNFS.Models;

namespace ItTech.Tool.AddonNFS.Utils
{
    /// <summary>
    /// Gerenciador centralizado de formulários do add-on
    /// Controla instâncias abertas e hierarquia mestre/filho
    /// </summary>
    public static class FormManager
    {
        // Dicionário para rastrear formulários abertos
        private static Dictionary<string, FormInfo> _formsAbertos = new Dictionary<string, FormInfo>();

        // Cache de dados dos grupos
        private static Dictionary<string, GrupoLote> _cacheGrupos = new Dictionary<string, GrupoLote>();

        // Lock para thread safety
        private static readonly object _lockForms = new object();
        private static readonly object _lockCache = new object();

        // IDs dos formulários
        public const string FORM_SELECIONAR_GRUPO = "ITTECH_NFS_SEL";
        public const string FORM_GRUPO_LOTE = "ITTECH_NFS_GRUPO";
        public const string FORM_VISUALIZACAO_LINHAS = "ITTECH_NFS_LINHAS";
        public const string FORM_RESULTADO = "ITTECH_NFS_RESULT";
        public const string FORM_CONFIG = "ITTECH_NFS_CFG";

        /// <summary>
        /// Informações sobre um formulário aberto
        /// </summary>
        private class FormInfo
        {
            public string FormUID { get; set; }
            public string TipoForm { get; set; }
            public bool IsMaster { get; set; }
            public string MasterFormUID { get; set; }
            public DateTime DataAbertura { get; set; }
            public string GrupoCode { get; set; }
            public int Etapa { get; set; }
            public SAPbouiCOM.Form FormReference { get; set; }
        }

        #region Cache de Grupos

        /// <summary>
        /// Adiciona ou atualiza um grupo no cache
        /// </summary>
        public static void CachearGrupo(string grupoCode, GrupoLote grupo)
        {
            lock (_lockCache)
            {
                if (!string.IsNullOrEmpty(grupoCode) && grupo != null)
                {
                    _cacheGrupos[grupoCode] = grupo;
                }
            }
        }

        /// <summary>
        /// Obtém um grupo do cache
        /// </summary>
        public static GrupoLote ObterGrupoCache(string grupoCode)
        {
            lock (_lockCache)
            {
                return _cacheGrupos.ContainsKey(grupoCode) ? _cacheGrupos[grupoCode] : null;
            }
        }

        /// <summary>
        /// Remove um grupo do cache
        /// </summary>
        public static void RemoverGrupoCache(string grupoCode)
        {
            lock (_lockCache)
            {
                if (_cacheGrupos.ContainsKey(grupoCode))
                {
                    _cacheGrupos.Remove(grupoCode);
                }
            }
        }

        /// <summary>
        /// Limpa todo o cache
        /// </summary>
        public static void LimparCache()
        {
            lock (_lockCache)
            {
                _cacheGrupos.Clear();
            }
        }

        #endregion

        #region Controle de Formulários

        /// <summary>
        /// Registra um formulário como aberto
        /// </summary>
        public static void RegistrarFormulario(string formUID, string tipoForm, bool isMaster = false,
            string masterFormUID = null, string grupoCode = null, int etapa = 0)
        {
            lock (_lockForms)
            {
                try
                {
                    SAPbouiCOM.Form formRef = null;
                    try
                    {
                        formRef = Application.SBO_Application.Forms.Item(formUID);
                    }
                    catch { }

                    if (!_formsAbertos.ContainsKey(formUID))
                    {
                        _formsAbertos.Add(formUID, new FormInfo
                        {
                            FormUID = formUID,
                            TipoForm = tipoForm,
                            IsMaster = isMaster,
                            MasterFormUID = masterFormUID,
                            DataAbertura = DateTime.Now,
                            GrupoCode = grupoCode,
                            Etapa = etapa,
                            FormReference = formRef
                        });

                        Application.SBO_Application.SetStatusBarMessage($"Formulário '{tipoForm}' registrado",
                            SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    else
                    {
                        // Atualizar informações se já existe
                        var formInfo = _formsAbertos[formUID];
                        formInfo.GrupoCode = grupoCode ?? formInfo.GrupoCode;
                        formInfo.Etapa = etapa > 0 ? etapa : formInfo.Etapa;
                        formInfo.FormReference = formRef ?? formInfo.FormReference;
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage($"Erro ao registrar formulário: {ex.Message}",
                        SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }

        /// <summary>
        /// Remove o registro de um formulário fechado
        /// </summary>
        public static void RemoverFormulario(string formUID)
        {
            lock (_lockForms)
            {
                try
                {
                    if (_formsAbertos.ContainsKey(formUID))
                    {
                        var formInfo = _formsAbertos[formUID];

                        // Se for um formulário mestre, notificar sobre filhos
                        if (formInfo.IsMaster)
                        {
                            var filhos = _formsAbertos.Values
                                .Where(f => f.MasterFormUID == formUID)
                                .Count();

                            if (filhos > 0)
                            {
                                Application.SBO_Application.SetStatusBarMessage(
                                    $"Formulário mestre fechado. {filhos} formulário(s) filho(s) ainda aberto(s)",
                                    SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            }
                        }

                        _formsAbertos.Remove(formUID);
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage($"Erro ao remover formulário: {ex.Message}",
                        SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }

        /// <summary>
        /// Verifica se um formulário está aberto
        /// </summary>
        public static bool FormularioEstaAberto(string formUID)
        {
            lock (_lockForms)
            {
                try
                {
                    if (_formsAbertos.ContainsKey(formUID))
                    {
                        // Verificar se realmente está aberto no SAP
                        try
                        {
                            var form = Application.SBO_Application.Forms.Item(formUID);
                            return true;
                        }
                        catch
                        {
                            // Se não conseguir acessar, remover do dicionário
                            _formsAbertos.Remove(formUID);
                            return false;
                        }
                    }
                    return false;
                }
                catch
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Obtém a referência de um formulário aberto
        /// </summary>
        public static SAPbouiCOM.Form ObterFormulario(string formUID)
        {
            lock (_lockForms)
            {
                try
                {
                    if (_formsAbertos.ContainsKey(formUID))
                    {
                        return Application.SBO_Application.Forms.Item(formUID);
                    }
                    return null;
                }
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Traz um formulário para frente
        /// </summary>
        public static void TrazerParaFrente(string formUID)
        {
            try
            {
                var form = Application.SBO_Application.Forms.Item(formUID);

                // Tornar visível
                if (!form.Visible)
                {
                    form.Visible = true;
                }

                // Restaurar se minimizado
                if (form.State == SAPbouiCOM.BoFormStateEnum.fs_Minimized)
                {
                    form.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                }

                // Selecionar (trazer para frente)
                form.Select();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao trazer formulário para frente: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// Fecha todos os formulários filhos de um formulário mestre
        /// </summary>
        public static void FecharFormulariosFilhos(string masterFormUID)
        {
            lock (_lockForms)
            {
                try
                {
                    var filhos = _formsAbertos.Values
                        .Where(f => f.MasterFormUID == masterFormUID)
                        .Select(f => f.FormUID)
                        .ToList();

                    foreach (var formUID in filhos)
                    {
                        try
                        {
                            var form = Application.SBO_Application.Forms.Item(formUID);
                            form.Close();
                        }
                        catch
                        {
                            // Formulário já foi fechado
                        }

                        // Remover do dicionário
                        _formsAbertos.Remove(formUID);
                    }

                    if (filhos.Count > 0)
                    {
                        Application.SBO_Application.SetStatusBarMessage($"{filhos.Count} formulário(s) filho(s) fechado(s)",
                            SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage($"Erro ao fechar formulários filhos: {ex.Message}",
                        SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }

        /// <summary>
        /// Verifica se existem formulários filhos abertos
        /// </summary>
        public static bool ExistemFormulariosFilhosAbertos(string masterFormUID = null)
        {
            lock (_lockForms)
            {
                if (string.IsNullOrEmpty(masterFormUID))
                {
                    return _formsAbertos.Values.Any(f => !f.IsMaster && !string.IsNullOrEmpty(f.MasterFormUID));
                }
                else
                {
                    return _formsAbertos.Values.Any(f => f.MasterFormUID == masterFormUID);
                }
            }
        }

        /// <summary>
        /// Fecha todos os formulários do add-on
        /// </summary>
        public static void FecharTodosFormularios()
        {
            lock (_lockForms)
            {
                try
                {
                    var formsParaFechar = _formsAbertos.Keys.ToList();

                    foreach (var formUID in formsParaFechar)
                    {
                        try
                        {
                            var form = Application.SBO_Application.Forms.Item(formUID);
                            form.Close();
                        }
                        catch
                        {
                            // Formulário já foi fechado
                        }
                    }

                    _formsAbertos.Clear();
                    LimparCache();

                    Application.SBO_Application.SetStatusBarMessage("Todos os formulários do add-on foram fechados",
                        SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage($"Erro ao fechar todos os formulários: {ex.Message}",
                        SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }

        #endregion

        #region Consultas

        /// <summary>
        /// Obtém o UID do formulário por grupo code
        /// </summary>
        public static string ObterFormularioPorGrupo(string grupoCode)
        {
            lock (_lockForms)
            {
                var form = _formsAbertos.Values.FirstOrDefault(f => f.GrupoCode == grupoCode);
                return form?.FormUID;
            }
        }

        /// <summary>
        /// Verifica se um grupo específico está sendo processado
        /// </summary>
        public static bool GrupoEstaEmProcessamento(string grupoCode)
        {
            lock (_lockForms)
            {
                return _formsAbertos.Values.Any(f => f.GrupoCode == grupoCode);
            }
        }

        /// <summary>
        /// Obtém o UID do formulário por tipo
        /// </summary>
        public static string ObterFormularioPorTipo(string tipoForm)
        {
            lock (_lockForms)
            {
                var form = _formsAbertos.Values.FirstOrDefault(f => f.TipoForm == tipoForm);
                return form?.FormUID;
            }
        }

        /// <summary>
        /// Obtém a etapa atual de um grupo
        /// </summary>
        public static int ObterEtapaAtualGrupo(string grupoCode)
        {
            lock (_lockForms)
            {
                var forms = _formsAbertos.Values
                    .Where(f => f.GrupoCode == grupoCode)
                    .OrderByDescending(f => f.Etapa);

                return forms.FirstOrDefault()?.Etapa ?? 0;
            }
        }

        /// <summary>
        /// Obtém informações detalhadas sobre formulários abertos
        /// </summary>
        public static string ObterStatusFormularios()
        {
            lock (_lockForms)
            {
                var status = $"Formulários abertos: {_formsAbertos.Count}\n";
                status += $"Grupos em cache: {_cacheGrupos.Count}\n";

                if (_formsAbertos.Count == 0)
                {
                    return status;
                }

                status += "━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n";

                // Primeiro mostrar formulários mestres
                var mestres = _formsAbertos.Values.Where(f => f.IsMaster).OrderBy(f => f.DataAbertura);
                foreach (var form in mestres)
                {
                    status += $"📋 {form.TipoForm}\n";
                    status += $"   ID: {form.FormUID}\n";
                    status += $"   Aberto em: {form.DataAbertura:dd/MM/yyyy HH:mm:ss}\n";

                    // Contar filhos
                    var filhos = _formsAbertos.Values.Where(f => f.MasterFormUID == form.FormUID).Count();
                    if (filhos > 0)
                    {
                        status += $"   Formulários filhos: {filhos}\n";
                    }
                    status += "\n";
                }

                // Depois mostrar formulários filhos
                var formFilhos = _formsAbertos.Values.Where(f => !f.IsMaster).OrderBy(f => f.DataAbertura);
                if (formFilhos.Any())
                {
                    status += "Formulários Filhos:\n";
                    foreach (var form in formFilhos)
                    {
                        status += $"  └─ {form.TipoForm}\n";
                        status += $"     ID: {form.FormUID}\n";
                        if (!string.IsNullOrEmpty(form.GrupoCode))
                        {
                            status += $"     Grupo: {form.GrupoCode}\n";
                            status += $"     Etapa: {form.Etapa}\n";
                        }
                        status += $"     Pai: {form.MasterFormUID}\n";
                        status += $"     Aberto em: {form.DataAbertura:HH:mm:ss}\n\n";
                    }
                }

                return status;
            }
        }

        #endregion
    }
}