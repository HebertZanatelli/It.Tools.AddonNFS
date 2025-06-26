using ItTech.Tool.AddonNFS.Controllers;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Utils;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using Application = SAPbouiCOM.Framework.Application;
using Company = SAPbobsCOM.Company;

namespace ItTech.Tool.AddonNFS.Forms
{
    [FormAttribute("ITTECH_NFS_SEL", "Forms/FormSelecionarGrupo.b1f")]
    class FormSelecionarGrupo : UserFormBase
    {
        #region Campos e Propriedades

        private GrupoLoteController _grupoController;
        private List<GrupoLote> _grupos;
        private int _ultimaLinhaSelecionada = -1;
        private bool _atualizandoDados = false;

        // Controles do formulário
        private SAPbouiCOM.Button BtnNovo;
        private SAPbouiCOM.Button BtnAtualizar;
        private SAPbouiCOM.Button BtnCancelar;
        private SAPbouiCOM.Button BtnAbrir;
        private SAPbouiCOM.Matrix Matrix;
        private SAPbouiCOM.StaticText StaticText0;
        // NOVO - Campos para detectar duplo clique
        private DateTime _ultimoClique = DateTime.MinValue;
        private int _ultimaLinhaClicada = -1;
        private const int INTERVALO_DUPLO_CLIQUE = 700; // milissegundos
        #endregion

        #region Construtor

        public FormSelecionarGrupo()
        {
        }

        #endregion

        #region Inicialização

        public override void OnInitializeComponent()
        {
            this.BtnNovo = ((SAPbouiCOM.Button)(this.GetItem("btnNovo").Specific));
            this.BtnAtualizar = ((SAPbouiCOM.Button)(this.GetItem("btnAtu").Specific));
            this.BtnCancelar = ((SAPbouiCOM.Button)(this.GetItem("btnCan").Specific));
            this.BtnAbrir = ((SAPbouiCOM.Button)(this.GetItem("btnAbrir").Specific));
            this.Matrix = ((SAPbouiCOM.Matrix)(this.GetItem("oMatrix").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));

            // Eventos dos botões
            this.BtnNovo.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnNovo_ClickBefore);
            this.BtnAtualizar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnAtualizar_ClickBefore);
            this.BtnCancelar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnCancelar_ClickBefore);
            this.BtnAbrir.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnAbrir_ClickBefore);
            ((SAPbouiCOM._IMatrixEvents_Event)this.Matrix).ClickBefore += Matrix_ClickBefore;

            this.OnCustomInitialize();
        }

        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);
            this.CloseBefore += new CloseBeforeHandler(this.Form_CloseBefore);
            this.ClickAfter += new ClickAfterHandler(this.Form_ClickAfter);
            this.KeyDownAfter += new KeyDownAfterHandler(this.Form_KeyDownAfter);
            //this.ClickBefore += new ClickBeforeHandler(this.Form_ClickBefore); // NOVO - Adicionar ClickBefore
        }

        private void Matrix_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Garante que o clique foi em linha válida
                if (pVal.Row > 0)
                {
                    // Força a seleção da linha
                    Matrix.SelectRow(pVal.Row, true, false);
                    _ultimaLinhaSelecionada = pVal.Row;

                    // Detecção de duplo clique
                    DateTime agora = DateTime.Now;
                    TimeSpan intervalo = agora - _ultimoClique;

                    bool isDuploClique = (pVal.Row == _ultimaLinhaClicada &&
                                          intervalo.TotalMilliseconds <= INTERVALO_DUPLO_CLIQUE);

                    if (isDuploClique)
                    {
                        AbrirGrupoSelecionado(pVal.Row - 1);
                        _ultimoClique = DateTime.MinValue;
                        _ultimaLinhaClicada = -1;
                    }
                    else
                    {
                        _ultimoClique = agora;
                        _ultimaLinhaClicada = pVal.Row;
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro Matrix_ClickBefore: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void OnCustomInitialize()
        {
            try
            {
                _grupoController = new GrupoLoteController((Company)Application.SBO_Application.Company.GetDICompany());

                // Criar DataTable se não existir
                if (!DataTableExists("dtGrupos"))
                {
                    UIAPIRawForm.DataSources.DataTables.Add("dtGrupos");
                }

                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtGrupos");
                if (dt.Columns.Count == 0)
                {
                    dt.Columns.Add("Code", BoFieldsType.ft_AlphaNumeric, 20);
                    dt.Columns.Add("Nome", BoFieldsType.ft_AlphaNumeric, 100);
                    dt.Columns.Add("DtLanc", BoFieldsType.ft_Date);
                    dt.Columns.Add("DtDoc", BoFieldsType.ft_Date);
                    dt.Columns.Add("Arquivo", BoFieldsType.ft_AlphaNumeric, 150);
                    dt.Columns.Add("Status", BoFieldsType.ft_AlphaNumeric, 50);
                    dt.Columns.Add("Total", BoFieldsType.ft_Integer);
                    dt.Columns.Add("Proc", BoFieldsType.ft_Integer);
                    dt.Columns.Add("Erro", BoFieldsType.ft_Integer);
                }

                ConfigurarMatrix();
                CarregarGrupos();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao inicializar: {ex.Message}", 1, "Ok", "", "");
            }
        }

        #endregion

        #region Configuração da Matrix

        private void ConfigurarMatrix()
        {
            try
            {

                // Vincular colunas ao DataTable
                Matrix.Columns.Item("Code").DataBind.Bind("dtGrupos", "Code");
                Matrix.Columns.Item("Nome").DataBind.Bind("dtGrupos", "Nome");
                Matrix.Columns.Item("DtLanc").DataBind.Bind("dtGrupos", "DtLanc");
                Matrix.Columns.Item("DtDoc").DataBind.Bind("dtGrupos", "DtDoc");
                Matrix.Columns.Item("Arquivo").DataBind.Bind("dtGrupos", "Arquivo");
                Matrix.Columns.Item("Status").DataBind.Bind("dtGrupos", "Status");
                Matrix.Columns.Item("Total").DataBind.Bind("dtGrupos", "Total");
                Matrix.Columns.Item("Proc").DataBind.Bind("dtGrupos", "Proc");
                Matrix.Columns.Item("Erro").DataBind.Bind("dtGrupos", "Erro");


            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao configurar matrix: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Carregamento de Dados

        private void CarregarGrupos()
        {
            try
            {
                _atualizandoDados = true;
                UIAPIRawForm.Freeze(true);

                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtGrupos");
                dt.Rows.Clear();

                _grupos = _grupoController.ListarGrupos();

                for (int i = 0; i < _grupos.Count; i++)
                {
                    var g = _grupos[i];
                    dt.Rows.Add();
                    dt.SetValue("Code", i, g.Code);
                    dt.SetValue("Nome", i, g.Nome);
                    dt.SetValue("DtLanc", i, g.DataLancamento);
                    dt.SetValue("DtDoc", i, g.DataDocumento);
                    dt.SetValue("Arquivo", i, g.NomeArquivo ?? "");
                    dt.SetValue("Status", i, ObterDescricaoStatus(g.Status));
                    dt.SetValue("Total", i, g.TotalLinhas);
                    dt.SetValue("Proc", i, g.LinhasProcessadas);
                    dt.SetValue("Erro", i, g.LinhasErro);
                }

                Matrix.Clear();
                Matrix.LoadFromDataSource();
                Matrix.AutoResizeColumns();

                // Selecionar primeira linha se houver grupos
                if (_grupos.Count > 0)
                {
                    Matrix.SelectRow(1, true, false);
                    _ultimaLinhaSelecionada = 1;
                }

                Application.SBO_Application.SetStatusBarMessage($"{_grupos.Count} grupos carregados",
                    BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao carregar grupos: {ex.Message}", 1, "Ok", "", "");
            }
            finally
            {
                _atualizandoDados = false;
                UIAPIRawForm.Freeze(false);
            }
        }

        #endregion

        #region Eventos do Formulário

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {
            try
            {
                UIAPIRawForm.Title = "Processar NFS-e em Lote - Selecionar Grupo";

                // Registrar no FormManager como formulário mestre
                FormManager.RegistrarFormulario(
                    UIAPIRawForm.UniqueID,
                    "Seleção de Grupo",
                    isMaster: true
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro no load: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void Form_CloseBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Verificar se há formulários filhos abertos
                if (FormManager.ExistemFormulariosFilhosAbertos(UIAPIRawForm.UniqueID))
                {
                    int resposta = Application.SBO_Application.MessageBox(
                        "Existem formulários de processamento abertos.\n" +
                        "Deseja fechá-los também?",
                        2, "Sim", "Não", "");

                    if (resposta == 1) // Sim
                    {
                        FormManager.FecharFormulariosFilhos(UIAPIRawForm.UniqueID);
                    }
                    else
                    {
                        BubbleEvent = false;
                        return;
                    }
                }

                // Remover do registro
                FormManager.RemoverFormulario(UIAPIRawForm.UniqueID);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao fechar: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Eventos de Interação

        private void Form_ClickBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Verificar se o clique foi na Matrix
                if (pVal.ItemUID == "oMatrix" && pVal.Row > 0)
                {
                    DateTime agora = DateTime.Now;
                    TimeSpan intervalo = agora - _ultimoClique;

                    bool isDuploClique = (pVal.Row == _ultimaLinhaClicada &&
                                         intervalo.TotalMilliseconds <= INTERVALO_DUPLO_CLIQUE);

                    if (isDuploClique && pVal.Row <= _grupos.Count)
                    {
                        AbrirGrupoSelecionado(pVal.Row - 1);
                        _ultimoClique = DateTime.MinValue;
                        _ultimaLinhaClicada = -1;
                    }
                    else
                    {
                        // Sempre seleciona a linha clicada
                        Matrix.SelectRow(pVal.Row, true, false);

                        _ultimaLinhaSelecionada = pVal.Row;
                        _ultimoClique = agora;
                        _ultimaLinhaClicada = pVal.Row;

                        if (pVal.Row <= _grupos.Count)
                        {
                            var grupo = _grupos[pVal.Row - 1];
                            Application.SBO_Application.SetStatusBarMessage(
                                $"Grupo selecionado: {grupo.Nome} - Status: {ObterDescricaoStatus(grupo.Status)}",
                                BoMessageTime.bmt_Short, false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro na seleção: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }



        private void Form_ClickAfter(SBOItemEventArg pVal)
        {
            try
            {
                if (_atualizandoDados) return;

                // Apenas manter controle se foi na Matrix
                if (pVal.ItemUID == "oMatrix" && pVal.Row > 0)
                {
                    _ultimaLinhaSelecionada = pVal.Row;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro no clique: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        // MANTER - DoubleClickAfter já está correto
        private void Form_DoubleClickAfter(SBOItemEventArg pVal)
        {
            try
            {
                // Verificar se o duplo clique foi na Matrix
                if (pVal.ItemUID == "oMatrix" && pVal.Row > 0 && pVal.Row <= _grupos.Count)
                {
                    // Abrir diretamente o grupo ao dar duplo clique
                    AbrirGrupoSelecionado(pVal.Row - 1);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro no duplo clique: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void Form_KeyDownAfter(SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ItemUID == "oMatrix")
                {
                    // Enter abre o grupo selecionado
                    if (pVal.CharPressed == 13) // Enter
                    {
                        int selectedRow = Matrix.GetNextSelectedRow(0, BoOrderType.ot_RowOrder);
                        if (selectedRow > 0 && selectedRow <= _grupos.Count)
                        {
                            AbrirGrupoSelecionado(selectedRow - 1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro no teclado: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Métodos de Navegação

        private void AbrirGrupoSelecionado(int index)
        {
            try
            {
                if (index < 0 || index >= _grupos.Count)
                {
                    Application.SBO_Application.MessageBox("Índice de grupo inválido", 1, "Ok", "", "");
                    return;
                }

                var grupoSelecionado = _grupos[index];

                // Verificar se este grupo já está sendo processado
                string formAberto = FormManager.ObterFormularioPorGrupo(grupoSelecionado.Code);
                if (!string.IsNullOrEmpty(formAberto))
                {
                    FormManager.TrazerParaFrente(formAberto);
                    Application.SBO_Application.SetStatusBarMessage(
                        "Este grupo já está aberto em outro formulário",
                        BoMessageTime.bmt_Short, false);
                    return;
                }

                Application.SBO_Application.SetStatusBarMessage($"Abrindo grupo: {grupoSelecionado.Nome}",
                    BoMessageTime.bmt_Short, false);

                // Obter o grupo completo para determinar a etapa
                var grupoCompleto = _grupoController.ObterGrupo(grupoSelecionado.Code);

                // Cachear o grupo
                FormManager.CachearGrupo(grupoSelecionado.Code, grupoCompleto);

                // Determinar qual formulário abrir baseado no status e conteúdo
                if (grupoCompleto.Status == StatusGrupo.Novo && grupoCompleto.TotalLinhas == 0)
                {
                    // Etapa 1: Grupo sem linhas importadas
                    AbrirFormularioGrupoLote(grupoSelecionado.Code);
                }
                else if (grupoCompleto.Status == StatusGrupo.Novo && grupoCompleto.TotalLinhas > 0)
                {
                    // Etapa 2: Tem linhas mas não processadas
                    AbrirFormularioVisualizacaoLinhas(grupoSelecionado.Code);
                }
                else if (grupoCompleto.Status != StatusGrupo.Novo)
                {
                    // Etapa 3: Já foi processado (parcial ou completo)
                    AbrirFormularioResultado(grupoSelecionado);
                }
                else
                {
                    // Caso não previsto, abrir na etapa 1
                    AbrirFormularioGrupoLote(grupoSelecionado.Code);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao abrir grupo: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void AbrirFormularioGrupoLote(string grupoCode)
        {
            try
            {
                //FormGrupoLote formGrupo = new FormGrupoLote(UIAPIRawForm.UniqueID);
                //// Chamar SetGrupoCode após Show()
                //formGrupo.SetGrupoCode(grupoCode);
                //formGrupo.Show();



                //FormManager.RegistrarFormulario(
                //    formGrupo.UIAPIRawForm.UniqueID,
                //    "Grupo Lote - Etapa 1",
                //    isMaster: false,
                //    masterFormUID: UIAPIRawForm.UniqueID,
                //    grupoCode: grupoCode,
                //    etapa: 1
                //);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao abrir formulário: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void AbrirFormularioVisualizacaoLinhas(string grupoCode)
        {
            try
            {
                FormVisualizacaoLinhas formLinhas = new FormVisualizacaoLinhas();
                formLinhas.Show();

                // Chamar SetCodeGroup após Show()
                formLinhas.SetCodeGroup(grupoCode);

                FormManager.RegistrarFormulario(
                    formLinhas.UIAPIRawForm.UniqueID,
                    "Visualização de Linhas - Etapa 2",
                    isMaster: false,
                    masterFormUID: UIAPIRawForm.UniqueID,
                    grupoCode: grupoCode,
                    etapa: 2
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao abrir formulário: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void AbrirFormularioResultado(GrupoLote grupo)
        {
            try
            {
                FormProcessamentoNFS formResultado = new FormProcessamentoNFS(UIAPIRawForm.UniqueID);

                formResultado.SetGrupo(grupo);
                formResultado.Show();



                FormManager.RegistrarFormulario(
                    formResultado.UIAPIRawForm.UniqueID,
                   "Resultado do Processamento - Etapa 3",
                    isMaster: false,
                    masterFormUID: UIAPIRawForm.UniqueID,
                    grupoCode: grupo.Code,
                    etapa: 3
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao abrir formulário: {ex.Message}", 1, "Ok", "", "");
            }
        }

        #endregion

        #region Eventos de Botões

        private void BtnNovo_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Verificar se já existe um formulário de criação aberto
                string formAberto = FormManager.ObterFormularioPorTipo("Grupo Lote - Etapa 1");
                if (!string.IsNullOrEmpty(formAberto) && FormManager.FormularioEstaAberto(formAberto))
                {
                    FormManager.TrazerParaFrente(formAberto);
                    Application.SBO_Application.SetStatusBarMessage("Formulário de criação já está aberto",
                        BoMessageTime.bmt_Short, false);
                }
                else
                {
                    // Abrir formulário de criação de grupo (Etapa 1)
                  //  FormGrupoLote formGrupo = new FormGrupoLote(UIAPIRawForm.UniqueID);
                    FormGrupoLote formGrupo = new FormGrupoLote();
                    formGrupo.Show();

                    FormManager.RegistrarFormulario(
                        formGrupo.UIAPIRawForm.UniqueID,
                        "Novo Grupo - Etapa 1",
                        isMaster: false,
                        masterFormUID: UIAPIRawForm.UniqueID,
                        etapa: 1
                    );
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void BtnAtualizar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Manter a linha selecionada ao atualizar
                int linhaSelecionada = Matrix.GetNextSelectedRow(0, BoOrderType.ot_RowOrder);

                CarregarGrupos();

                // Restaurar seleção se possível
                if (linhaSelecionada > 0 && linhaSelecionada <= Matrix.RowCount)
                {
                    Matrix.SelectRow(linhaSelecionada, true, false);
                    _ultimaLinhaSelecionada = linhaSelecionada;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao atualizar: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void BtnCancelar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            UIAPIRawForm.Close();
        }

        // MODIFICAR - BtnAbrir para usar o método auxiliar
        private void BtnAbrir_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                int indice = ObterIndiceSelecionado();

                if (indice < 0)
                {
                    Application.SBO_Application.MessageBox("Por favor, selecione um grupo para abrir.", 1, "Ok", "", "");
                    return;
                }

                AbrirGrupoSelecionado(indice);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao abrir grupo: {ex.Message}", 1, "Ok", "", "");
            }
        }

        #endregion

        #region Métodos Auxiliares
        // ADICIONAR - Método auxiliar para garantir que temos o índice correto
        private int ObterIndiceSelecionado()
        {
            try
            {
                int selectedRow = Matrix.GetNextSelectedRow(0, BoOrderType.ot_RowOrder);
                if (selectedRow > 0)
                {
                    return selectedRow - 1; // Converter para índice baseado em zero
                }

                // Se não houver seleção formal, usar a última linha clicada
                if (_ultimaLinhaSelecionada > 0 && _ultimaLinhaSelecionada <= _grupos.Count)
                {
                    return _ultimaLinhaSelecionada - 1;
                }

                return -1;
            }
            catch
            {
                return -1;
            }
        }
        private bool DataTableExists(string id)
        {
            for (int i = 0; i < UIAPIRawForm.DataSources.DataTables.Count; i++)
            {
                if (UIAPIRawForm.DataSources.DataTables.Item(i).UniqueID == id)
                    return true;
            }
            return false;
        }

        private string ObterDescricaoStatus(StatusGrupo status)
        {
            switch (status)
            {
                case StatusGrupo.Novo: return "Novo";
                case StatusGrupo.EmProcessamento: return "Em Processamento";
                case StatusGrupo.ProcessadoCompleto: return "Processado Completo";
                case StatusGrupo.ProcessadoParcial: return "Processado Parcial";
                case StatusGrupo.Erro: return "Erro";
                default: return "Desconhecido";
            }
        }

        #endregion
    }
}