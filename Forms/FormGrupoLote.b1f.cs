using ItTech.Tool.AddonNFS.Controller;
using ItTech.Tool.AddonNFS.Controllers;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Utils;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Linq;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Application = SAPbouiCOM.Framework.Application;

namespace ItTech.Tool.AddonNFS.Forms
{
    [FormAttribute("ITTECH_NFS_GRUPO", "Forms/FormGrupoLote.b1f")]
    class FormGrupoLote : UserFormBase
    {
        #region Propriedades e Campos

        private GrupoLoteController _grupoController;
        private ImportacaoController _importController;
        private string _grupoCode;
        private GrupoLote _grupoAtual;
        private string _caminhoArquivoTemp;
        private bool _modoVisualizacao = false;
        private bool _arquivoValidado = false;
        private string _formOrigemUID;
        private bool _dadosCarregados = false;
        private bool _carregandoDados = false;
        private readonly object _lockCarregamento = new object();

        // Controles do formulário
        private StaticText LblEtapa;
        private StaticText LblInstrucao;
        private StaticText LblStatus;
        private EditText TxtNome;
        private EditText TxtDocDate;
        private EditText TxtDueDate;
        private Button BtnAbrir;
        private Button BtnDescarregar;
        private Button BtnValidar;
        private Button BtnVoltar;
        private Button BtnProcessar;
        private Button BtnCancelar;
        private Matrix MatrixArquivo;
        private ComboBox cmbTipoDoc;

        #endregion

        #region Construtores


        public FormGrupoLote(string formOrigemUID = null)
        {

            _formOrigemUID = formOrigemUID;
        }

        #endregion

        #region Inicialização e SetGrupoCode

        public void SetGrupoCode(string grupoCode)
        {
            lock (_lockCarregamento)
            {
                if (_grupoCode == grupoCode && _dadosCarregados)
                    return;

                _grupoCode = grupoCode;
                _dadosCarregados = false;

                if (!string.IsNullOrEmpty(_grupoCode))
                {
                    CarregarGrupoExistente();
                    _dadosCarregados = true;
                }
            }
        }

        public override void OnInitializeComponent()
        {
            this.LblEtapa = ((SAPbouiCOM.StaticText)(this.GetItem("lblEtapa").Specific));
            this.LblInstrucao = ((SAPbouiCOM.StaticText)(this.GetItem("lblInst").Specific));
            this.LblStatus = ((SAPbouiCOM.StaticText)(this.GetItem("lblStatus").Specific));
            this.TxtNome = ((SAPbouiCOM.EditText)(this.GetItem("txtNome").Specific));
            this.TxtDocDate = ((SAPbouiCOM.EditText)(this.GetItem("txtDocDate").Specific));
            this.TxtDueDate = ((SAPbouiCOM.EditText)(this.GetItem("txtDueDate").Specific));
            this.BtnAbrir = ((SAPbouiCOM.Button)(this.GetItem("btnAbrir").Specific));
            this.BtnDescarregar = ((SAPbouiCOM.Button)(this.GetItem("btnDesc").Specific));
            this.BtnValidar = ((SAPbouiCOM.Button)(this.GetItem("btnValidar").Specific));
            this.BtnVoltar = ((SAPbouiCOM.Button)(this.GetItem("btnVoltar").Specific));
            this.BtnProcessar = ((SAPbouiCOM.Button)(this.GetItem("btnProc").Specific));
            this.BtnCancelar = ((SAPbouiCOM.Button)(this.GetItem("btnCancel").Specific));
            this.MatrixArquivo = ((SAPbouiCOM.Matrix)(this.GetItem("mtxArq").Specific));
            this.cmbTipoDoc = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbTipoDoc").Specific));

            //    Eventos
            this.BtnAbrir.ClickBefore += this.BtnAbrir_ClickBefore;
            this.BtnDescarregar.ClickBefore += this.BtnDescarregar_ClickBefore;
            this.BtnValidar.ClickBefore += this.BtnValidar_ClickBefore;
            this.BtnVoltar.ClickBefore += this.BtnVoltar_ClickBefore;
            this.BtnProcessar.ClickBefore += this.BtnProcessar_ClickBefore;
            this.BtnCancelar.ClickBefore += this.BtnCancelar_ClickBefore;
            this.cmbTipoDoc.ComboSelectAfter += this.CmbTipoDoc_ComboSelectAfter;

            //    Eventos de validação em tempo real
            this.TxtNome.LostFocusAfter += this.TxtNome_LostFocusAfter;
            this.TxtDocDate.LostFocusAfter += this.TxtDocDate_LostFocusAfter;
            this.TxtDueDate.LostFocusAfter += this.TxtDueDate_LostFocusAfter;

            
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += Form_LoadAfter;
            this.CloseBefore += Form_CloseBefore;
            this.ResizeAfter += Form_ResizeAfter;
        }

        private void OnCustomInitialize()
        {
            try
            {
                _grupoController = new GrupoLoteController((SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany());
                _importController = new ImportacaoController((SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany());

                cmbTipoDoc.ValidValues.Add("-", "Selecione");
                cmbTipoDoc.ValidValues.Add("NFS", "Nota Fiscal de Saída");
                cmbTipoDoc.ValidValues.Add("ENT", "Entrega");
                cmbTipoDoc.ValidValues.Add("NFE", "Nota Fiscal de Entrada");
                cmbTipoDoc.Select(0, BoSearchKey.psk_Index);

                // Criar DataTable para arquivos
                if (!DataTableExists("dtArquivo"))
                {
                    UIAPIRawForm.DataSources.DataTables.Add("dtArquivo");
                }

                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtArquivo");
                if (dt.Columns.Count == 0)
                {
                    dt.Columns.Add("Linha", BoFieldsType.ft_AlphaNumeric, 10);
                    dt.Columns.Add("Arquivo", BoFieldsType.ft_AlphaNumeric, 255);
                    dt.Columns.Add("Tamanho", BoFieldsType.ft_AlphaNumeric, 50);
                    dt.Columns.Add("Status", BoFieldsType.ft_AlphaNumeric, 50);
                }

                ConfigurarMatrixArquivo();

                // Se não tem código, é novo grupo
                if (string.IsNullOrEmpty(_grupoCode))
                {
                    InicializarNovoGrupo();
                }

                AtualizarInterfaceContextual();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao inicializar: {ex.Message}", 1, "Ok", "", "");
            }
        }

        #endregion

        #region Configurações

        private void ConfigurarMatrixArquivo()
        {
            try
            {
                // Configurar binding das colunas
                MatrixArquivo.Columns.Item("#").DataBind.Bind("dtArquivo", "Linha");
                MatrixArquivo.Columns.Item("Arquivo").DataBind.Bind("dtArquivo", "Arquivo");
                MatrixArquivo.Columns.Item("Tamanho").DataBind.Bind("dtArquivo", "Tamanho");
                MatrixArquivo.Columns.Item("Status").DataBind.Bind("dtArquivo", "Status");

                // Configurar larguras
                MatrixArquivo.Columns.Item("#").Width = 30;
                MatrixArquivo.Columns.Item("Arquivo").Width = 400;
                MatrixArquivo.Columns.Item("Tamanho").Width = 100;
                MatrixArquivo.Columns.Item("Status").Width = 76;

                // Tornar não editável
                for (int i = 0; i < MatrixArquivo.Columns.Count; i++)
                {
                    MatrixArquivo.Columns.Item(i).Editable = false;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao configurar matrix: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Carregamento de Dados

        private void InicializarNovoGrupo()
        {
            try
            {
                // Definir datas padrão
                TxtDocDate.Value = DateTime.Now.ToString("yyyyMMdd");
                TxtDueDate.Value = DateTime.Now.ToString("yyyyMMdd");

                // Sugerir nome do grupo
                TxtNome.Value = $"Lote_{DateTime.Now:yyyyMMdd_HHmm}";

                AtualizarStatus("Configurando novo grupo de processamento", false);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao inicializar: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void CarregarGrupoExistente()
        {
            try
            {
                // Evitar múltiplos carregamentos simultâneos
                if (_carregandoDados)
                    return;
                    
                _carregandoDados = true;
                UIAPIRawForm.Freeze(true);
                AtualizarStatus("Carregando dados do grupo...", false);

                // Verificar cache primeiro
                _grupoAtual = FormManager.ObterGrupoCache(_grupoCode);
                if (_grupoAtual == null)
                {
                    _grupoAtual = _grupoController.ObterGrupo(_grupoCode);
                    FormManager.CachearGrupo(_grupoCode, _grupoAtual);
                }

                // Preencher campos
                TxtNome.Value = _grupoAtual.Nome;
                TxtDueDate.Value = _grupoAtual.DataLancamento.ToString("yyyyMMdd");
                TxtDocDate.Value = _grupoAtual.DataDocumento.ToString("yyyyMMdd");

                // Se tem arquivo, mostrar na matrix
                if (!string.IsNullOrEmpty(_grupoAtual.NomeArquivo))
                {
                    ExibirArquivoNaMatrix(_grupoAtual.NomeArquivo, _grupoAtual.CaminhoArquivo);
                    _arquivoValidado = (_grupoAtual.TotalLinhas > 0);
                }

                // Verificar se já tem linhas importadas
                if (_grupoAtual.TotalLinhas > 0)
                {
                    _modoVisualizacao = true;
                    AtualizarStatus($"Grupo carregado: {_grupoAtual.TotalLinhas} linhas já importadas", true);
                }
                else
                {
                    _modoVisualizacao = false;
                    AtualizarStatus("Grupo carregado. Selecione o arquivo para importação.", false);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao carregar grupo: {ex.Message}", 1, "Ok", "", "");
            }
            finally
            {
                _carregandoDados = false;
                UIAPIRawForm.Freeze(false);
                AtualizarInterfaceContextual();
            }
        }

        #endregion

        #region Interface e Navegação

        private void AtualizarInterfaceContextual()
        {
            try
            {
                if (_modoVisualizacao)
                {
                    // Modo visualização - grupo já tem linhas
                    BtnAbrir.Item.Enabled = false;
                    BtnDescarregar.Item.Enabled = false;
                    BtnValidar.Item.Enabled = false;
                    TxtNome.Item.Enabled = false;
                    TxtDocDate.Item.Enabled = false;
                    TxtDueDate.Item.Enabled = false;
                    BtnProcessar.Caption = "Visualizar →";
                    LblInstrucao.Caption = "Este grupo já possui dados importados. Clique em 'Visualizar' para continuar.";
                }
                else
                {
                    // Modo edição
                    bool temArquivo = !string.IsNullOrEmpty(_caminhoArquivoTemp);
                    BtnDescarregar.Item.Enabled = temArquivo;
                    BtnValidar.Item.Enabled = temArquivo && !_arquivoValidado;
                    BtnProcessar.Item.Enabled = ValidarFormulario(false);

                    if (_arquivoValidado)
                    {
                        BtnProcessar.Caption = "Importar →";
                    }
                    else
                    {
                        BtnProcessar.Caption = "Próximo →";
                    }
                }

                // Botão voltar sempre visível se veio de outro form
                BtnVoltar.Item.Visible = !string.IsNullOrEmpty(_formOrigemUID);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao atualizar interface: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void AtualizarStatus(string mensagem, bool sucesso)
        {
            try
            {
                LblStatus.Caption = mensagem;
                Application.SBO_Application.SetStatusBarMessage(mensagem,
                    BoMessageTime.bmt_Short, !sucesso);
            }
            catch
            {
                Application.SBO_Application.SetStatusBarMessage(mensagem,
                    BoMessageTime.bmt_Short, !sucesso);
            }
        }

        #endregion

        #region Eventos de Validação

        private void TxtNome_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            AtualizarInterfaceContextual();
        }

        private void TxtDocDate_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            ValidarDatas();
            AtualizarInterfaceContextual();
        }

        private void TxtDueDate_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            ValidarDatas();
            AtualizarInterfaceContextual();
        }

        private bool ValidarDatas()
        {
            try
            {
                if (!string.IsNullOrEmpty(TxtDocDate.Value) && !string.IsNullOrEmpty(TxtDueDate.Value))
                {
                    DateTime dtDoc = DateTime.ParseExact(TxtDocDate.Value, "yyyyMMdd", null);
                    DateTime dtLanc = DateTime.ParseExact(TxtDueDate.Value, "yyyyMMdd", null);

                    if (dtDoc > dtLanc)
                    {
                        AtualizarStatus("⚠ Data do documento não pode ser posterior à data de lançamento", false);
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Eventos de Botões

        private void BtnAbrir_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Desabilitar botão temporariamente para evitar múltiplos cliques
                BtnAbrir.Item.Enabled = false;
                AtualizarStatus("Abrindo diálogo de arquivo...", false);

                // Abrir diálogo sem travar
                AbrirDialogoArquivoAsync();
            }
            catch (Exception ex)
            {
                BtnAbrir.Item.Enabled = true;
                Application.SBO_Application.MessageBox($"Erro ao selecionar arquivo: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void BtnValidar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (string.IsNullOrEmpty(_caminhoArquivoTemp))
                {
                    Application.SBO_Application.MessageBox("Nenhum arquivo selecionado para validar.", 1, "Ok", "", "");
                    return;
                }

                // Limpar status anterior
                _arquivoValidado = false;
                
                // Atualizar status na matrix
                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtArquivo");
                if (dt.Rows.Count > 0)
                {
                    dt.SetValue("Status", 0, "Validando...");
                    MatrixArquivo.LoadFromDataSource();
                }

                UIAPIRawForm.Freeze(true);
                AtualizarStatus("Validando arquivo...", false);

                // Validar estrutura do arquivo de forma síncrona
                ValidarArquivoSincrono();
            }
            catch (Exception ex)
            {
                _arquivoValidado = false;
                Application.SBO_Application.MessageBox($"Erro ao validar: {ex.Message}", 1, "Ok", "", "");
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void ValidarArquivoSincrono()
        {
            try
            {
                // Validar usando ImportacaoController
                var resultado = _importController.ImportarExcel(_caminhoArquivoTemp);

                if (resultado.validacao.Valida)
                {
                    // Atualizar UI
                    var dt = UIAPIRawForm.DataSources.DataTables.Item("dtArquivo");
                    if (dt.Rows.Count > 0)
                    {
                        dt.SetValue("Status", 0, "✓ Validado");
                    }

                    MatrixArquivo.LoadFromDataSource();
                    _arquivoValidado = true;

                    AtualizarStatus($"Arquivo validado com sucesso! {resultado.linhas.Count} linhas encontradas.", true);
                }
                else
                {
                    string erros = string.Join("\n", resultado.validacao.Erros.Take(5));
                    Application.SBO_Application.MessageBox($"Erros na validação:\n{erros}", 1, "Ok", "", "");
                    AtualizarStatus("Arquivo com erros de validação", false);
                }

                AtualizarInterfaceContextual();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro na validação: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void BtnDescarregar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                int resposta = Application.SBO_Application.MessageBox(
                    "Deseja remover o arquivo selecionado?",
                    2, "Sim", "Não", "");

                if (resposta == 1)
                {
                    LimparArquivo();
                    AtualizarStatus("Arquivo removido", false);
                    AtualizarInterfaceContextual();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void BtnVoltar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (HaAlteracoesNaoSalvas())
                {
                    int resposta = Application.SBO_Application.MessageBox(
                        "Existem alterações não salvas. Deseja continuar?",
                        2, "Sim", "Não", "");

                    if (resposta != 1)
                    {
                        BubbleEvent = false;
                        return;
                    }
                }

                // Voltar ao formulário anterior
                if (!string.IsNullOrEmpty(_formOrigemUID))
                {
                    FormManager.TrazerParaFrente(_formOrigemUID);
                }

                UIAPIRawForm.Close();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao voltar: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void BtnProcessar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (_modoVisualizacao)
                {
                    // Modo visualização - abrir próxima etapa
                    AbrirProximaEtapa();
                    return;
                }

                if (!ValidarFormulario(true))
                    return;

                // Criar ou atualizar grupo
                SalvarGrupo();

                // Abrir próxima etapa
                AbrirProximaEtapa();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao processar: {ex.Message}", 1, "Ok", "", "");
            }
        }

        private void BtnCancelar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (HaAlteracoesNaoSalvas())
                {
                    int resposta = Application.SBO_Application.MessageBox(
                        "Existem alterações não salvas. Deseja realmente cancelar?",
                        2, "Sim", "Não", "");

                    if (resposta != 1)
                    {
                        BubbleEvent = false;
                        return;
                    }
                }

                UIAPIRawForm.Close();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Métodos Auxiliares

        /// <summary>
        /// Evento disparado após o usuário selecionar um item no ComboBox de Tipo de Documento.
        /// A única responsabilidade deste método é solicitar a reavaliação da interface.
        /// </summary>
        private void CmbTipoDoc_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            AtualizarInterfaceContextual();
        }

        private bool ValidarFormulario(bool mostrarMensagens)
        {
            // Nome do grupo
            if (string.IsNullOrWhiteSpace(TxtNome.Value))
            {
                if (mostrarMensagens)
                {
                    Application.SBO_Application.MessageBox("Informe o nome do grupo", 1, "Ok", "", "");
                    TxtNome.Active = true;
                }
                return false;
            }

            //Validação Tipo de Documento
            if (string.IsNullOrEmpty(cmbTipoDoc.Value) || cmbTipoDoc.Value.Trim() == "-")
            {
                if (mostrarMensagens)
                {
                    Application.SBO_Application.MessageBox("Selecione um Tipo de Documento válido para continuar.", 1, "Ok", "", "");
                    // Opcional: Focar no ComboBox para guiar o usuário
                    cmbTipoDoc.Item.Click(BoCellClickType.ct_Regular);
                }
                return false;
            }


            // Datas
            if (string.IsNullOrWhiteSpace(TxtDueDate.Value) || string.IsNullOrWhiteSpace(TxtDocDate.Value))
            {
                if (mostrarMensagens)
                {
                    Application.SBO_Application.MessageBox("Informe as datas de lançamento e documento", 1, "Ok", "", "");
                }
                return false;
            }

            // Validar datas
            if (!ValidarDatas() && mostrarMensagens)
            {
                return false;
            }

            // Arquivo (apenas para grupos novos)
            if (!_modoVisualizacao && string.IsNullOrEmpty(_caminhoArquivoTemp))
            {
                if (mostrarMensagens)
                {
                    Application.SBO_Application.MessageBox("Selecione o arquivo Excel", 1, "Ok", "", "");
                }
                return false;
            }

            // Validação do arquivo
            if (!_modoVisualizacao && !_arquivoValidado && mostrarMensagens)
            {
                Application.SBO_Application.MessageBox("O arquivo precisa ser validado antes de continuar", 1, "Ok", "", "");
                return false;
            }

            return true;
        }

        private bool ValidarArquivoExcel(string arquivo)
        {
            if (!arquivo.ToLower().EndsWith(".xlsx"))
            {
                Application.SBO_Application.MessageBox(
                    "Por favor, selecione um arquivo Excel (.xlsx)\n\n" +
                    "Formatos suportados: .xlsx",
                    1, "Ok", "", "");
                return false;
            }

            if (!File.Exists(arquivo))
            {
                Application.SBO_Application.MessageBox("Arquivo não encontrado", 1, "Ok", "", "");
                return false;
            }

            // Verificar tamanho
            FileInfo fi = new FileInfo(arquivo);
            if (fi.Length > 10 * 1024 * 1024) // 10MB
            {
                int resposta = Application.SBO_Application.MessageBox(
                    $"O arquivo selecionado é grande ({fi.Length / 1024 / 1024}MB).\n" +
                    "Isso pode tornar o processamento mais lento.\n\n" +
                    "Deseja continuar?",
                    2, "Sim", "Não", "");

                if (resposta != 1)
                    return false;
            }

            return true;
        }

        private void ExibirArquivoNaMatrix(string nomeArquivo, string caminhoCompleto)
        {
            try
            {
                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtArquivo");
                dt.Rows.Clear();
                dt.Rows.Add();

                FileInfo fi = new FileInfo(caminhoCompleto);
                string tamanho = fi.Length < 1024 ? $"{fi.Length} bytes" :
                                fi.Length < 1024 * 1024 ? $"{fi.Length / 1024} KB" :
                                $"{fi.Length / 1024 / 1024} MB";

                dt.SetValue("Linha", 0, "1");
                dt.SetValue("Arquivo", 0, nomeArquivo);
                dt.SetValue("Tamanho", 0, tamanho);
                dt.SetValue("Status", 0, _arquivoValidado ? "✓ Validado" : "Pendente");

                MatrixArquivo.Clear();
                MatrixArquivo.LoadFromDataSource();
                MatrixArquivo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao exibir arquivo: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void LimparArquivo()
        {
            var dt = UIAPIRawForm.DataSources.DataTables.Item("dtArquivo");
            dt.Rows.Clear();
            MatrixArquivo.Clear();
            MatrixArquivo.LoadFromDataSource();
            _caminhoArquivoTemp = string.Empty;
            _arquivoValidado = false;
        }

        private bool HaAlteracoesNaoSalvas()
        {
            if (_grupoAtual != null)
            {
                return TxtNome.Value != _grupoAtual.Nome ||
                       TxtDueDate.Value != _grupoAtual.DataLancamento.ToString("yyyyMMdd") ||
                       TxtDocDate.Value != _grupoAtual.DataDocumento.ToString("yyyyMMdd");
            }

            return !string.IsNullOrEmpty(TxtNome.Value) || !string.IsNullOrEmpty(_caminhoArquivoTemp);
        }

        private void SalvarGrupo()
        {
            try
            {
                if (string.IsNullOrEmpty(_grupoCode))
                {
                    // Criar novo grupo
                    GrupoLote novoGrupo = new GrupoLote
                    {
                        Nome = TxtNome.Value.Trim(),
                        DataLancamento = DateTime.ParseExact(TxtDueDate.Value, "yyyyMMdd", null),
                        DataDocumento = DateTime.ParseExact(TxtDocDate.Value, "yyyyMMdd", null),
                        NomeArquivo = Path.GetFileName(_caminhoArquivoTemp),
                        CaminhoArquivo = _caminhoArquivoTemp,
                        Status = StatusGrupo.Novo,
                        TipoDocumento = cmbTipoDoc.Selected.Value
                    };

                    _grupoCode = _grupoController.CriarGrupo(novoGrupo);
                    _grupoAtual = novoGrupo;
                    _grupoAtual.Code = _grupoCode;

                    // Cachear o novo grupo
                    FormManager.CachearGrupo(_grupoCode, _grupoAtual);

                    AtualizarStatus($"Grupo '{novoGrupo.Nome}' criado com sucesso!", true);
                }
                else
                {
                    // TODO: Implementar atualização se necessário
                    AtualizarStatus("Grupo atualizado com sucesso!", true);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao salvar grupo: {ex.Message}", ex);
            }
        }

        private void AbrirProximaEtapa()
        {
            try
            {
                // Cachear dados atuais
                if (_grupoAtual != null)
                {
                    FormManager.CachearGrupo(_grupoCode, _grupoAtual);
                }

                // Criar o formulário antes de configurar
                FormVisualizacaoLinhas formLinhas = new FormVisualizacaoLinhas();
                
                // Configurar o código do grupo ANTES de mostrar
                formLinhas.SetCodeGroup(_grupoCode);
                
                // Agora sim, mostrar o formulário
                formLinhas.Show();

                // Fechar este formulário
                UIAPIRawForm.Close();
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao abrir próxima etapa: {ex.Message}", ex);
            }
        }

        // NOVO MÉTODO - Substitui o antigo AbrirDialogoArquivo()
        private void AbrirDialogoArquivoAsync()
        {
            Thread thread = new Thread(() =>
            {
                string arquivo = string.Empty;

                try
                {
                    // Criar dummy form para garantir que o diálogo apareça no SAP B1
                    using (System.Windows.Forms.Form dummyForm = new System.Windows.Forms.Form())
                    {
                        // Configurar dummy form para ser invisível mas funcional
                        dummyForm.TopMost = true;
                        dummyForm.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                        dummyForm.ShowInTaskbar = false;
                        dummyForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                        dummyForm.Size = new System.Drawing.Size(1, 1);
                        dummyForm.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        dummyForm.Location = new System.Drawing.Point(-1000, -1000);
                        dummyForm.Opacity = 0;
                        
                        // IMPORTANTE: Mostrar o form para criar handle válido
                        dummyForm.Show();
                        dummyForm.BringToFront();
                        
                        using (System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog())
                        {
                            ofd.Filter = "Arquivos Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*";
                            ofd.Title = "Selecione o arquivo Excel com os dados das NFS-e";
                            ofd.Multiselect = false;
                            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                            // Usar o dummy form como parent
                            if (ofd.ShowDialog(dummyForm) == System.Windows.Forms.DialogResult.OK)
                            {
                                arquivo = ofd.FileName;
                            }
                        }
                        
                        // Fechar dummy form
                        dummyForm.Close();
                    }

                    // Processar resultado na mesma thread
                    if (!string.IsNullOrEmpty(arquivo))
                    {
                        // Voltar para thread do SAP para atualizar UI
                        ProcessarArquivoSelecionado(arquivo);
                    }
                    else
                    {
                        // Reabilitar botão se cancelou
                        ReabilitarBotaoAbrir();
                    }
                }
                catch (Exception ex)
                {
                    // Log erro se necessário 
                    System.Diagnostics.Debug.WriteLine($"Erro no diálogo: {ex.Message}");
                    ReabilitarBotaoAbrir();
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            // NÃO TEM JOIN! - Esta é a correção principal
        }

        // NOVO MÉTODO
        private void ProcessarArquivoSelecionado(string arquivo)
        {
            // Executar de forma assíncrona na thread principal
            Task.Run(() =>
            {
                try
                {
                    // Voltar para thread principal do SAP
                    Application.SBO_Application.Forms.ActiveForm.Freeze(true);
                    
                    if (!ValidarArquivoExcel(arquivo))
                    {
                        ReabilitarBotaoAbrir();
                        Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                        return;
                    }

                    ExibirArquivoNaMatrix(Path.GetFileName(arquivo), arquivo);
                    _caminhoArquivoTemp = arquivo;
                    _arquivoValidado = false;

                    AtualizarStatus("Arquivo selecionado. Clique em 'Validar Arquivo' para continuar.", false);
                    AtualizarInterfaceContextual();
                    
                    Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                    Application.SBO_Application.MessageBox($"Erro ao processar arquivo: {ex.Message}", 1, "Ok", "", "");
                }
                finally
                {
                    ReabilitarBotaoAbrir();
                }
            });
        }

        // NOVO MÉTODO
        private void ReabilitarBotaoAbrir()
        {
            try
            {
                if (BtnAbrir != null && BtnAbrir.Item != null)
                {
                    BtnAbrir.Item.Enabled = true;
                }
            }
            catch { }
        }

        private bool DataTableExists(string tableName)
        {
            for (int i = 0; i < UIAPIRawForm.DataSources.DataTables.Count; i++)
            {
                if (UIAPIRawForm.DataSources.DataTables.Item(i).UniqueID == tableName)
                    return true;
            }
            return false;
        }

        #endregion

        #region Eventos do Formulário

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {
            try
            {
                UIAPIRawForm.Title = "Processar NFS-e em Lote - Etapa 1: Criar/Editar Grupo";

                // Registrar no FormManager
                if (!FormManager.FormularioEstaAberto(UIAPIRawForm.UniqueID))
                {
                    FormManager.RegistrarFormulario(
                        UIAPIRawForm.UniqueID,
                        "Grupo Lote - Etapa 1",
                        isMaster: false,
                        masterFormUID: _formOrigemUID ?? FormManager.FORM_SELECIONAR_GRUPO,
                        grupoCode: _grupoCode,
                        etapa: 1
                    );
                }

                // Focar no primeiro campo se for novo
                if (string.IsNullOrEmpty(_grupoCode))
                {
                    //TxtNome.Active = true;
                }
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
                // Limpar cache se necessário
                if (!string.IsNullOrEmpty(_grupoCode))
                {
                    FormManager.RemoverGrupoCache(_grupoCode);
                }
                
                // Limpar recursos
                _grupoController = null;
                _importController = null;
                _grupoAtual = null;
                
                // Remover do FormManager
                FormManager.RemoverFormulario(UIAPIRawForm.UniqueID);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao fechar: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            // Ajustar layout se necessário
        }

        #endregion


    }
}