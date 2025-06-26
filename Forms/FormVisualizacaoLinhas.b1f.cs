using ItTech.Tool.AddonNFS.Controller;
using ItTech.Tool.AddonNFS.Controllers;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Utils;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using Application = SAPbouiCOM.Framework.Application;
using ItTech.Tool.AddonNFS.Services;

namespace ItTech.Tool.AddonNFS.Forms
{
    [FormAttribute("ITTECH_NFS_LINHAS", "Forms/FormVisualizacaoLinhas.b1f")]
    class FormVisualizacaoLinhas : UserFormBase
    {
        #region Propriedades e Campos

        private string _grupoCode;
        private string _formOrigemUID;
        private GrupoLote _grupo;
        private List<LinhaImportacao> _linhas;
        private ImportacaoController _importController;
        private GrupoLoteController _grupoController;
        private ServiceLayerInvoiceClient _serviceLayerClient;
        private ProcessamentoNFSController _processamentoController;

        // Estado
        private bool _dadosCarregados = false;
        private readonly object _lockCarregamento = new object();
        private CancellationTokenSource _cancellationTokenSource;
        private bool _carregandoDados = false;
        private int _totalLinhasSelecionadas = 0;
        private decimal _valorTotalSelecionado = 0;
        private bool _atualizandoSelecao = false;
        private bool _estadoCheckboxSelAll = false;
        private DateTime _ultimaAtualizacaoContadores = DateTime.MinValue; // NOVO

        // Cache para evitar recálculos
        private Dictionary<int, LinhaImportacao> _linhasIndexadas;
        private HashSet<int> _linhasSelecionadasIndex;
        private string _ultimoFiltroXml = "";

        // Controles do formulário
        private StaticText LblEtapa;
        private StaticText LblInstrucao;
        private StaticText LblContadores;
        private StaticText LblStatusImport;
        private EditText TxtGrupo;
        private EditText TxtDtLanc;
        private EditText TxtDtDoc;
        private ComboBox CmbFiltro;
        private Matrix MtxLinhas;
        private Button BtnVoltar;
        private Button BtnValidar;
        private Button BtnProcessar;
        private Button BtnCancelar;
        private CheckBox CheckBox0;

        // Thread-safe UI updates
        private System.Timers.Timer _timerUI;
        private Queue<Action> _filaAcoesUI = new Queue<Action>();
        private readonly object _lockFila = new object();
        private bool _formFechando = false;

        // Constantes para otimização
        private const int BATCH_SIZE = 100;
        private const int XML_THRESHOLD = 50;

        #endregion

        #region Construtores
     
        public FormVisualizacaoLinhas() : this(null, null) { }

        public FormVisualizacaoLinhas(string grupoCode, string formOrigemUID = null)
        {
            _grupoCode = grupoCode;
            _formOrigemUID = formOrigemUID;
            _linhasIndexadas = new Dictionary<int, LinhaImportacao>();
            _linhasSelecionadasIndex = new HashSet<int>();
        }

        #endregion

        #region SetCodeGroup com Thread Segura

        public void SetCodeGroup(string code)
        {
            lock (_lockCarregamento)
            {
                if (_grupoCode == code && _dadosCarregados)
                    return;

                _grupoCode = code;
                _dadosCarregados = false;

                // Cancelar operações anteriores se existirem
                _cancellationTokenSource?.Cancel();
                _cancellationTokenSource = new CancellationTokenSource();

                if (!string.IsNullOrEmpty(_grupoCode))
                {
                    // Iniciar timer para processar ações UI
                    IniciarTimerUI();

                    // Thread para carregamento em segundo plano
                    Thread threadCarregamento = new Thread(() =>
                    {
                        var token = _cancellationTokenSource.Token;
                        try
                        {
                            // Mostrar progresso
                            EnfileirarAcaoUI(() =>
                            {
                                Application.SBO_Application.SetStatusBarMessage(
                                    "Carregando dados do grupo...",
                                    BoMessageTime.bmt_Short,
                                    false);
                            });

                            // Verificar cancelamento
                            if (token.IsCancellationRequested) return;

                            CarregarDadosAsync(token);
                            _dadosCarregados = true;

                            // Finalizar
                            EnfileirarAcaoUI(() =>
                            {
                                Application.SBO_Application.SetStatusBarMessage(
                                    "Dados carregados com sucesso!",
                                    BoMessageTime.bmt_Short,
                                    false);

                                // Parar timer quando terminar
                                PararTimerUI();
                            });
                        }
                        catch (Exception ex)
                        {
                            EnfileirarAcaoUI(() =>
                            {
                                if (!_formFechando && !token.IsCancellationRequested)
                                {
                                    Application.SBO_Application.SetStatusBarMessage(
                                        $"Erro ao carregar dados: {ex.Message}",
                                        BoMessageTime.bmt_Short,
                                        true);
                                }
                                PararTimerUI();
                            });
                        }
                        finally
                        {
                            // Garantir que ProgressBar seja fechada
                            ProgressBarHelper.Instance.Fechar();
                        }
                    });

                    threadCarregamento.IsBackground = true;
                    threadCarregamento.Priority = ThreadPriority.BelowNormal;
                    threadCarregamento.Start();
                }
            }
        }

        private void IniciarTimerUI()
        {
            if (_timerUI == null)
            {
                _timerUI = new System.Timers.Timer(50);
                _timerUI.Elapsed += ProcessarFilaUI;
                _timerUI.Start();
            }
        }

        private void PararTimerUI()
        {
            if (_timerUI != null && _filaAcoesUI.Count == 0)
            {
                _timerUI.Stop();
                _timerUI.Dispose();
                _timerUI = null;
                
                // Limpar fila restante
                lock (_lockFila)
                {
                    _filaAcoesUI.Clear();
                }
            }
        }

        private void EnfileirarAcaoUI(Action acao)
        {
            lock (_lockFila)
            {
                _filaAcoesUI.Enqueue(acao);
            }
        }

        private void ProcessarFilaUI(object sender, System.Timers.ElapsedEventArgs e)
        {
            List<Action> acoesParaProcessar = new List<Action>();

            lock (_lockFila)
            {
                // OTIMIZADO: Pegar até 3 ações por vez
                int maxAcoes = Math.Min(3, _filaAcoesUI.Count);
                for (int i = 0; i < maxAcoes; i++)
                {
                    if (_filaAcoesUI.Count > 0)
                        acoesParaProcessar.Add(_filaAcoesUI.Dequeue());
                }
            }

            foreach (var acao in acoesParaProcessar)
            {
                try
                {
                    if (!_formFechando)
                    {
                        acao.Invoke();
                    }
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage(
                        $"Erro UI: {ex.Message}",
                        BoMessageTime.bmt_Short,
                        true);
                }
            }
        }

        private void CarregarDadosAsync(CancellationToken cancellationToken)
        {
            try
            {
                if (string.IsNullOrEmpty(_grupoCode))
                    return;

                _carregandoDados = true;

                // Usar ProgressBarHelper
                ProgressBarHelper.Instance.Criar("Carregando dados do grupo...", 100);

                // Verificar cancelamento
                if (cancellationToken.IsCancellationRequested)
                {
                    ProgressBarHelper.Instance.Fechar();
                    return;
                }

                // Fase 1: Carregar grupo
                ProgressBarHelper.Instance.Atualizar(20, "Obtendo dados do grupo...");

                _grupo = FormManager.ObterGrupoCache(_grupoCode);
                if (_grupo == null)
                {
                    _grupo = _grupoController.ObterGrupo(_grupoCode);
                    FormManager.CachearGrupo(_grupoCode, _grupo);
                }

                // Fase 2: Atualizar UI com informações do cabeçalho
                ProgressBarHelper.Instance.Atualizar(40, "Atualizando interface...");

                EnfileirarAcaoUI(() =>
                {
                    try
                    {
                        UIAPIRawForm.Freeze(true);
                        MostrarProgressoCarregamento(true);

                        TxtGrupo.Value = _grupo.Nome;
                        TxtDtLanc.Value = _grupo.DataLancamento.ToString("dd/MM/yyyy");
                        TxtDtDoc.Value = _grupo.DataDocumento.ToString("dd/MM/yyyy");
                    }
                    finally
                    {
                        UIAPIRawForm.Freeze(false);
                    }
                });

                // Verificar cancelamento
                if (cancellationToken.IsCancellationRequested) return;

                // Fase 3: Verificar se precisa importar ou já tem dados
                ProgressBarHelper.Instance.Atualizar(60, "Verificando dados...");

                if (_grupo.Linhas == null || _grupo.Linhas.Count == 0)
                {
                    ImportarDadosAsync(cancellationToken);
                }
                else
                {
                    _linhas = _grupo.Linhas;
                    IndexarLinhas();

                    // Fase 4: Carregar linhas na UI
                    ProgressBarHelper.Instance.Atualizar(80, "Carregando linhas na tela...");

                    EnfileirarAcaoUI(() =>
                    {
                        try
                        {
                            UIAPIRawForm.Freeze(true);
                            CarregarLinhasNaMatrixOtimizado();
                            AtualizarContadores();
                            AtualizarInterface();
                            MostrarProgressoCarregamento(false);
                        }
                        finally
                        {
                            UIAPIRawForm.Freeze(false);
                        }
                    });

                    ProgressBarHelper.Instance.Atualizar(100, "Concluído!");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao carregar dados: {ex.Message}");
            }
            finally
            {
                _carregandoDados = false;
                ProgressBarHelper.Instance.Fechar();
            }
        }

        private void ImportarDadosAsync(CancellationToken cancellationToken)
        {
            try
            {
                // Progress bar já foi criada em CarregarDadosAsync
                ProgressBarHelper.Instance.Atualizar(65, "Importando dados do arquivo Excel...");

                var resultado = _importController.ImportarExcelComProgresso(
               _grupo.CaminhoArquivo,
               (processadas, msg) =>
               {
                   int progresso = 65 + Math.Min(10, (processadas * 10 / Math.Max(processadas, 100)));
                   ProgressBarHelper.Instance.Atualizar(progresso, msg);

                   // Verificar cancelamento
                   if (cancellationToken.IsCancellationRequested)
                       throw new OperationCanceledException();
               }
                 );

                if (!resultado.validacao.Valida)
                {
                    throw new Exception(string.Join("\n", resultado.validacao.Erros.Take(5)));
                }

                // Verificar cancelamento
                if (cancellationToken.IsCancellationRequested) return;

                // Atualizar progresso
                ProgressBarHelper.Instance.Atualizar(75, $"Salvando {resultado.linhas.Count} linhas no banco...");

                _importController.SalvarLinhasComProgresso(_grupoCode, resultado.linhas, (processadas, msg) =>
                {
                    int progresso = 75 + (processadas * 10 / resultado.linhas.Count);
                    ProgressBarHelper.Instance.Atualizar(progresso, msg);
                });

                // Atualizar total de linhas no grupo
                _grupoController.AtualizarTotalLinhas(_grupoCode, resultado.linhas.Count);

                ProgressBarHelper.Instance.Atualizar(85, "Recarregando dados...");

                // Recarregar as linhas do banco para obter os Codes corretos
                _linhas = _grupoController.ObterLinhasGrupo(_grupoCode);

                // Indexar linhas para acesso rápido
                IndexarLinhas();

                // Atualizar cache
                _grupo.Linhas = _linhas;
                _grupo.TotalLinhas = _linhas.Count;
                FormManager.CachearGrupo(_grupoCode, _grupo);

                ProgressBarHelper.Instance.Atualizar(95, "Finalizando...");

                // Carregar na Matrix via UI thread
                EnfileirarAcaoUI(() =>
                {
                    try
                    {
                        UIAPIRawForm.Freeze(true);
                        CarregarLinhasNaMatrixOtimizado();
                        AtualizarContadores();
                        AtualizarInterface();
                        MostrarProgressoCarregamento(false);
                    }
                    finally
                    {
                        UIAPIRawForm.Freeze(false);
                    }
                });

                ProgressBarHelper.Instance.Atualizar(100, "Importação concluída!");
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro na importação: {ex.Message}");
            }
        }

        #endregion

        #region Inicialização

        public override void OnInitializeComponent()
        {
            //    Labels e textos
            this.LblEtapa = ((SAPbouiCOM.StaticText)(this.GetItem("lblEtapa").Specific));
            this.LblInstrucao = ((SAPbouiCOM.StaticText)(this.GetItem("lblInst").Specific));
            this.LblContadores = ((SAPbouiCOM.StaticText)(this.GetItem("lblCont").Specific));
            this.LblStatusImport = ((SAPbouiCOM.StaticText)(this.GetItem("lblStImp").Specific));
            //    Campos
            this.TxtGrupo = ((SAPbouiCOM.EditText)(this.GetItem("txtGrp").Specific));
            this.TxtDtLanc = ((SAPbouiCOM.EditText)(this.GetItem("txtDtLan").Specific));
            this.TxtDtDoc = ((SAPbouiCOM.EditText)(this.GetItem("txtDtDoc").Specific));
            //    Controles de seleção
            this.CmbFiltro = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbFiltro").Specific));
            this.CmbFiltro.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CmbFiltro_ComboSelectAfter);
            //    Matrix
            this.MtxLinhas = ((SAPbouiCOM.Matrix)(this.GetItem("mtxLin").Specific));
            //    Botões
            this.BtnVoltar = ((SAPbouiCOM.Button)(this.GetItem("btnVoltar").Specific));
            this.BtnValidar = ((SAPbouiCOM.Button)(this.GetItem("btnValid").Specific));
            this.BtnProcessar = ((SAPbouiCOM.Button)(this.GetItem("btnProc").Specific));
            this.BtnProcessar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnProcessar_ClickBefore);
            this.BtnCancelar = ((SAPbouiCOM.Button)(this.GetItem("btnCanc").Specific));
            this.BtnVoltar.ClickBefore += this.BtnVoltar_ClickBefore;
            this.BtnValidar.ClickBefore += this.BtnValidar_ClickBefore;
            this.BtnProcessar.ClickBefore += this.BtnProcessar_ClickBefore;
            this.BtnCancelar.ClickBefore += this.BtnCancelar_ClickBefore;
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("chkAll").Specific));
            this.CheckBox0.PressedBefore += this.ClickCheckAfter;
            this.OnCustomInitialize();

        }

        private void OnCustomInitialize()
        {
            try
            {
                var company = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                _importController = new ImportacaoController(company);
                _grupoController = new GrupoLoteController(company);
                _serviceLayerClient = new ServiceLayerInvoiceClient();
                _processamentoController = new ProcessamentoNFSController(company, _serviceLayerClient);

                if (_linhasIndexadas == null)
                    _linhasIndexadas = new Dictionary<int, LinhaImportacao>();
                if (_linhasSelecionadasIndex == null)
                    _linhasSelecionadasIndex = new HashSet<int>();

                // Define altura fixa do formulário
                UIAPIRawForm.Height = 600;

                //// Define altura máxima para a Matrix


                // Criar DataTable otimizado
                ConfigurarDataTable();

                // Configurar Matrix
                ConfigurarMatrix();

                // Configurar filtros
                CmbFiltro.Select(0, BoSearchKey.psk_Index);
                this.CheckBox0.Checked = true;

                // Configurar navegação
                BtnVoltar.Item.Visible = !string.IsNullOrEmpty(_formOrigemUID);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao inicializar: {ex.Message}", 1, "Ok", "", "");
            }
        }

        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += Form_LoadAfter;
            this.CloseBefore += Form_CloseBefore;
            this.ResizeAfter += Form_ResizeAfter;
        }

        #endregion

        #region Configuração DataTable e Matrix

        private void ConfigurarDataTable()
        {
            try
            {
                if (!DataTableExists("dtLinhas"))
                {
                    UIAPIRawForm.DataSources.DataTables.Add("dtLinhas");
                }

                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");

                if (dt.Columns.Count == 0)
                {
                    // Adicionar todas as colunas de uma vez
                    var colunas = new[]
                    {
                        ("Sel", BoFieldsType.ft_AlphaNumeric, 1),
                        ("Linha", BoFieldsType.ft_Integer, 0),
                        ("Filial", BoFieldsType.ft_AlphaNumeric, 10),
                        ("CodCli", BoFieldsType.ft_AlphaNumeric, 15),
                        ("NomeCli", BoFieldsType.ft_AlphaNumeric, 100),
                        ("CodItem", BoFieldsType.ft_AlphaNumeric, 20),
                        ("DescItem", BoFieldsType.ft_AlphaNumeric, 100),
                        ("Utiliz", BoFieldsType.ft_AlphaNumeric, 10),
                        ("CodImp", BoFieldsType.ft_AlphaNumeric, 10),
                        ("SeqNF", BoFieldsType.ft_AlphaNumeric, 10),
                        ("CondPagto", BoFieldsType.ft_AlphaNumeric, 10),
                        ("Valor", BoFieldsType.ft_Float, 0),
                        ("ObsNF", BoFieldsType.ft_AlphaNumeric, 254),
                        ("TipoTrib", BoFieldsType.ft_AlphaNumeric, 50),
                        ("Status", BoFieldsType.ft_AlphaNumeric, 20),
                        ("Code", BoFieldsType.ft_AlphaNumeric, 50)
                    };

                    foreach (var (nome, tipo, tamanho) in colunas)
                    {
                        if (tamanho > 0)
                            dt.Columns.Add(nome, tipo, tamanho);
                        else
                            dt.Columns.Add(nome, tipo);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao configurar DataTable: {ex.Message}");
            }
        }

        private void ConfigurarMatrix()
        {
            try
            {

                // Definir configurações das colunas
                var colunas = new[]
                {
                   // ("#", BoFormItemTypes.it_EDIT, 30, false, ""),
                    ("Sel", BoFormItemTypes.it_CHECK_BOX, 35, true, "✓"),
                    ("Linha", BoFormItemTypes.it_EDIT, 35, false, "#"),
                    ("Filial", BoFormItemTypes.it_EDIT, 35, false, "Filial"),
                    ("CodCli", BoFormItemTypes.it_LINKED_BUTTON, 120, false, "Código Cliente"),
                    ("NomeCli", BoFormItemTypes.it_EDIT, 200, false, "Nome Cliente"),
                    ("CodItem", BoFormItemTypes.it_LINKED_BUTTON, 100, false, "Código Item"),
                    ("DescItem", BoFormItemTypes.it_EDIT, 150, false, "Descrição"),
                    ("Utiliz", BoFormItemTypes.it_EDIT, 70, false, "Utilização"),
                    ("CodImp", BoFormItemTypes.it_EDIT, 80, false, "Imposto"),
                    ("SeqNF", BoFormItemTypes.it_EDIT, 60, false, "Seq NF"),
                    ("CondPagto", BoFormItemTypes.it_EDIT, 80, false, "Cond.Pgto"),
                    ("Valor", BoFormItemTypes.it_EDIT, 100, false, "Valor"),
                    ("ObsNF", BoFormItemTypes.it_EDIT, 200, false, "Observação NF"),
                    ("TipoTrib", BoFormItemTypes.it_EDIT, 100, false, "Tipo Trib"),
                    ("Status", BoFormItemTypes.it_EDIT, 80, false, "Status")
                };

                // Criar todas as colunas
                foreach (var (id, tipo, largura, editavel, titulo) in colunas)
                {
                    var col = MtxLinhas.Columns.Item(id);
                    col.Width = largura;
                    col.Editable = editavel;

                    if (!string.IsNullOrEmpty(titulo))
                        col.TitleObject.Caption = titulo;

                    if (id != "#")
                        col.DataBind.Bind("dtLinhas", id);
                }

                // Configurar checkbox
                var colSel = MtxLinhas.Columns.Item("Sel");
                colSel.ValOn = "Y";
                colSel.ValOff = "N";

                // Configurar LinkedButtons
                var lnkCliente = (LinkedButton)MtxLinhas.Columns.Item("CodCli").ExtendedObject;
                lnkCliente.LinkedObject = BoLinkedObject.lf_BusinessPartner;

                var lnkItem = (LinkedButton)MtxLinhas.Columns.Item("CodItem").ExtendedObject;
                lnkItem.LinkedObject = BoLinkedObject.lf_Items;

                // Configurar eventos da Matrix
                MtxLinhas.ClickAfter += Matrix_ClickAfter;
                MtxLinhas.ValidateAfter += Matrix_ValidateAfter;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao configurar Matrix: {ex.Message}");
            }
        }

        #endregion

        #region Carregamento de Dados Otimizado

        private void IndexarLinhas()
        {
            _linhasIndexadas.Clear();
            for (int i = 0; i < _linhas.Count; i++)
            {
                _linhasIndexadas[i] = _linhas[i];
            }
        }

        private void CarregarLinhasNaMatrixOtimizado()
        {
            try
            {
                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");

                // Se houver muitas linhas, usar processamento em lote
                if (_linhas.Count > BATCH_SIZE)
                {
                    CarregarLinhasEmLote(dt);
                }
                else
                {
                    CarregarLinhasSimples(dt);
                }

                MtxLinhas.Clear();
                MtxLinhas.LoadFromDataSource();
                MtxLinhas.AutoResizeColumns();

                // Selecionar todas por padrão se nenhuma estiver selecionada
                if (_linhas.Count > 0 && _linhas.All(l => !l.Selecionada))
                {
                    _estadoCheckboxSelAll = true;
                    SelecionarTodasLinhasOtimizado(true);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao carregar linhas: {ex.Message}");
            }
        }

        private void CarregarLinhasSimples(DataTable dt)
        {
            dt.Rows.Clear();

            // OTIMIZADO: Processar em mini-lotes mesmo para arquivos pequenos
            int miniLoteSize = 25;

            for (int i = 0; i < _linhas.Count; i += miniLoteSize)
            {
                int end = Math.Min(i + miniLoteSize, _linhas.Count);

                for (int j = i; j < end; j++)
                {
                    var linha = _linhas[j];
                    dt.Rows.Add();
                    PreencherLinhaDataTable(dt, j, linha);
                }
            }
        }

        private void CarregarLinhasEmLote(DataTable dt)
        {
            dt.Rows.Clear();

            // Criar progress bar para carregamento em lote
            ProgressBarHelper.Instance.ExecutarComProgress(
                $"Carregando {_linhas.Count} linhas...",
                _linhas.Count,
                (atualizador) =>
                {
                    for (int batch = 0; batch < _linhas.Count; batch += BATCH_SIZE)
                    {
                        int end = Math.Min(batch + BATCH_SIZE, _linhas.Count);

                        for (int i = batch; i < end; i++)
                        {
                            var linha = _linhas[i];
                            dt.Rows.Add();
                            PreencherLinhaDataTable(dt, i, linha);
                        }

                        // Atualizar progress bar
                        atualizador(end, $"Carregadas {end} de {_linhas.Count} linhas...");
                    }
                });
        }

        private void PreencherLinhaDataTable(DataTable dt, int index, LinhaImportacao linha)
        {
            dt.SetValue("Sel", index, linha.Selecionada ? "Y" : "N");
            dt.SetValue("Linha", index, linha.NumeroLinha);
            dt.SetValue("Filial", index, linha.Filial ?? "");
            dt.SetValue("CodCli", index, linha.CodigoCliente ?? "");
            dt.SetValue("NomeCli", index, linha.NomeCliente ?? "");
            dt.SetValue("CodItem", index, linha.CodigoItem ?? "");
            dt.SetValue("DescItem", index, linha.DescricaoItem ?? "");
            dt.SetValue("Utiliz", index, linha.Utilizacao ?? "");
            dt.SetValue("CodImp", index, linha.CodigoImposto ?? "");
            dt.SetValue("SeqNF", index, linha.CodSeq ?? "");
            dt.SetValue("CondPagto", index, linha.CondicaoPagamento ?? "");
            dt.SetValue("Valor", index, Convert.ToDouble(linha.Valor));
            dt.SetValue("ObsNF", index, linha.ObservacaoNF ?? "");
            dt.SetValue("TipoTrib", index, linha.TipoTributacao ?? "");
            dt.SetValue("Status", index, ObterDescricaoStatus(linha.Status));
            dt.SetValue("Code", index, linha.Code ?? "");
        }

        #endregion

        #region Eventos de Seleção Otimizados

        private void ClickCheckAfter(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Se já está no estado desejado, não fazer nada
                if (_estadoCheckboxSelAll == this.CheckBox0.Checked)
                    return;

                var marcarTodos = this.CheckBox0.Checked;
                UIAPIRawForm.Freeze(true);

                _estadoCheckboxSelAll = marcarTodos;
                SelecionarTodasLinhasOtimizado(marcarTodos);
                AtualizarContadoresOtimizado();
                AtualizarInterface();

                Application.SBO_Application.SetStatusBarMessage(
                    marcarTodos ? "Todas as linhas foram selecionadas" : "Todas as linhas foram desmarcadas",
                    BoMessageTime.bmt_Short,
                    false
                );
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao selecionar/desmarcar linhas: {ex.Message}", 1, "Ok", "", "");
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void SelecionarTodasLinhasOtimizado(bool selecionar)
        {
            try
            {
                _atualizandoSelecao = true;
                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");

                // Se houver muitas linhas, mostrar progress bar
                if (dt.Rows.Count > 500)
                {
                    ProgressBarHelper.Instance.ExecutarComProgress(
                        selecionar ? "Selecionando todas as linhas..." : "Desmarcando todas as linhas...",
                        100,
                        (atualizador) =>
                        {
                            atualizador(20, "Preparando seleção...");

                            // Otimização para grandes volumes usando XML
                            if (dt.Rows.Count > XML_THRESHOLD)
                            {
                                atualizador(50, "Processando seleção via XML...");
                                SelecionarViaXML(dt, selecionar);
                            }
                            else
                            {
                                atualizador(50, "Processando seleção...");
                                SelecionarViaLoop(dt, selecionar);
                            }

                            atualizador(80, "Atualizando interface...");
                            MtxLinhas.LoadFromDataSource();

                            // Atualizar modelo e índices
                            _linhasSelecionadasIndex.Clear();
                            if (_linhas != null)
                            {
                                for (int i = 0; i < _linhas.Count; i++)
                                {
                                    _linhas[i].Selecionada = selecionar;
                                    if (selecionar)
                                        _linhasSelecionadasIndex.Add(i);
                                }
                            }

                            atualizador(100, "Concluído!");
                        });
                }
                else
                {
                    // Para poucas linhas, fazer sem progress bar
                    if (dt.Rows.Count > XML_THRESHOLD)
                    {
                        SelecionarViaXML(dt, selecionar);
                    }
                    else
                    {
                        SelecionarViaLoop(dt, selecionar);
                    }

                    MtxLinhas.LoadFromDataSource();

                    // Atualizar modelo e índices
                    _linhasSelecionadasIndex.Clear();
                    if (_linhas != null)
                    {
                        for (int i = 0; i < _linhas.Count; i++)
                        {
                            _linhas[i].Selecionada = selecionar;
                            if (selecionar)
                                _linhasSelecionadasIndex.Add(i);
                        }
                    }
                }
            }
            finally
            {
                _atualizandoSelecao = false;
            }
        }

        private void SelecionarViaXML(DataTable dt, bool selecionar)
        {
            string xmlData = dt.SerializeAsXML(BoDataTableXmlSelect.dxs_DataOnly);

            if (selecionar)
            {
                xmlData = xmlData.Replace("<Cell><ColumnUid>Sel</ColumnUid><Value>N</Value></Cell>",
                                        "<Cell><ColumnUid>Sel</ColumnUid><Value>Y</Value></Cell>");
            }
            else
            {
                xmlData = xmlData.Replace("<Cell><ColumnUid>Sel</ColumnUid><Value>Y</Value></Cell>",
                                        "<Cell><ColumnUid>Sel</ColumnUid><Value>N</Value></Cell>");
            }

            dt.LoadSerializedXML(BoDataTableXmlSelect.dxs_DataOnly, xmlData);
        }

        private void SelecionarViaLoop(DataTable dt, bool selecionar)
        {
            string valor = selecionar ? "Y" : "N";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.SetValue("Sel", i, valor);
            }
        }

        /// <summary>
        /// ✅ TASK 2 FIX: Adiciona seleção visual da linha ao clicar em qualquer célula
        /// </summary>
        private void Matrix_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                // ✅ CORREÇÃO TASK 2: Sempre selecionar a linha visualmente ao clicar
                if (pVal.Row > 0)
                {
                    // Selecionar a linha visualmente - NOVA FUNCIONALIDADE
                    MtxLinhas.SelectRow(pVal.Row, true, false);
                }

                // LÓGICA ORIGINAL PRESERVADA: Tratamento especial para coluna "Sel"
                if (pVal.ColUID == "Sel" && pVal.Row > 0)
                {
                    var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");
                    var linha = pVal.Row - 1;
                    var valor = dt.GetValue("Sel", linha).ToString();
                    bool novoValor = valor != "Y";

                    dt.SetValue("Sel", linha, novoValor ? "Y" : "N");

                    // Atualizar índices
                    if (novoValor)
                        _linhasSelecionadasIndex.Add(linha);
                    else
                        _linhasSelecionadasIndex.Remove(linha);

                    // Atualizar modelo
                    if (_linhasIndexadas.ContainsKey(linha))
                        _linhasIndexadas[linha].Selecionada = novoValor;

                    AtualizarContadoresOtimizado();
                }
            }
            catch (Exception ex)
            {
               // Application.SBO_Application.SetStatusBarMessage($"Erro: {ex.Message}", BoMessageTime.bmt_Short, true);
            }
        }

        private void Matrix_ValidateAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == "Sel" && pVal.ItemChanged)
                {
                    AtualizarContadoresOtimizado();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro: {ex.Message}", BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Filtros e Contadores Otimizados

        private void CmbFiltro_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                UIAPIRawForm.Freeze(true);

                if (_linhas == null || _linhas.Count == 0)
                    return;

                int filtro = Convert.ToInt32(CmbFiltro.Selected.Value);

                // Usar cache se o mesmo filtro for aplicado
                string filtroKey = $"{filtro}_{_linhas.Count}";
                if (filtroKey == _ultimoFiltroXml && filtro == 0)
                {
                    UIAPIRawForm.Freeze(false);
                    return;
                }

                _ultimoFiltroXml = filtroKey;

                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");
                List<LinhaImportacao> linhasFiltradas;

                switch (filtro)
                {
                    case 1: // Apenas Selecionados
                        linhasFiltradas = _linhasSelecionadasIndex
                            .Select(i => _linhasIndexadas[i])
                            .ToList();
                        break;
                    case 2: // Não Selecionados
                        linhasFiltradas = _linhasIndexadas
                            .Where(kv => !_linhasSelecionadasIndex.Contains(kv.Key))
                            .Select(kv => kv.Value)
                            .ToList();
                        break;
                    default: // Mostrar Todos
                        linhasFiltradas = _linhas;
                        break;
                }

                // Recarregar DataTable otimizado
                dt.Rows.Clear();

                if (linhasFiltradas.Count > BATCH_SIZE)
                {
                    CarregarLinhasFiltradasEmLote(dt, linhasFiltradas);
                }
                else
                {
                    CarregarLinhasFiltradas(dt, linhasFiltradas);
                }

                MtxLinhas.Clear();
                MtxLinhas.LoadFromDataSource();
                MtxLinhas.AutoResizeColumns();

                AtualizarContadoresOtimizado();
                AtualizarInterface();

                Application.SBO_Application.SetStatusBarMessage("Filtro aplicado com sucesso.",
                    BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro no filtro: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void CarregarLinhasFiltradas(DataTable dt, List<LinhaImportacao> linhasFiltradas)
        {
            for (int i = 0; i < linhasFiltradas.Count; i++)
            {
                var linha = linhasFiltradas[i];
                dt.Rows.Add();
                PreencherLinhaDataTable(dt, i, linha);
            }
        }

        private void CarregarLinhasFiltradasEmLote(DataTable dt, List<LinhaImportacao> linhasFiltradas)
        {
            for (int batch = 0; batch < linhasFiltradas.Count; batch += BATCH_SIZE)
            {
                int end = Math.Min(batch + BATCH_SIZE, linhasFiltradas.Count);

                for (int i = batch; i < end; i++)
                {
                    var linha = linhasFiltradas[i];
                    dt.Rows.Add();
                    PreencherLinhaDataTable(dt, i, linha);
                }
            }
        }

        private void AtualizarContadores()
        {
            AtualizarContadoresOtimizado();
        }

        private void AtualizarContadoresOtimizado()
        {
            try
            {
                // OTIMIZADO: Evitar atualizações muito frequentes
                if ((DateTime.Now - _ultimaAtualizacaoContadores).TotalMilliseconds < 100)
                    return;

                _ultimaAtualizacaoContadores = DateTime.Now;

                if (_linhasSelecionadasIndex != null && _linhasSelecionadasIndex.Count > 0)
                {
                    _totalLinhasSelecionadas = _linhasSelecionadasIndex.Count;
                    _valorTotalSelecionado = _linhasSelecionadasIndex
                        .Where(i => _linhasIndexadas.ContainsKey(i))
                        .Sum(i => _linhasIndexadas[i].Valor);
                }
                else
                {
                    // Fallback para cálculo direto se índices não estiverem disponíveis
                    var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");
                    _totalLinhasSelecionadas = 0;
                    _valorTotalSelecionado = 0;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.GetValue("Sel", i).ToString() == "Y")
                        {
                            _totalLinhasSelecionadas++;
                            _valorTotalSelecionado += Convert.ToDecimal(dt.GetValue("Valor", i));
                        }
                    }
                }

                int totalLinhas = _linhas?.Count ?? 0;

                // OTIMIZADO: Atualizar label apenas se mudou
                string novoTexto = $"{_totalLinhasSelecionadas} de {totalLinhas} linhas selecionadas | " +
                                  $"Valor total: {_valorTotalSelecionado:C}";

                if (LblContadores.Caption != novoTexto)
                {
                    LblContadores.Caption = novoTexto;
                }

                BtnProcessar.Item.Enabled = _totalLinhasSelecionadas > 0;
                BtnValidar.Item.Enabled = _totalLinhasSelecionadas > 0;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao atualizar contadores: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Validação e Processamento

        private void BtnValidar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                var linhasSelecionadas = ObterLinhasSelecionadasOtimizado();

                if (linhasSelecionadas.Count == 0)
                {
                    Application.SBO_Application.MessageBox("Selecione pelo menos uma linha para validar.", 1, "Ok", "", "");
                    return;
                }

                UIAPIRawForm.Freeze(true);

                // Usar ProgressBarHelper para validação
                ProgressBarHelper.Instance.ExecutarComProgress(
                    $"Validando {linhasSelecionadas.Count} linhas...",
                    100,
                    (atualizador) =>
                    {
                        atualizador(20, "Preparando validação...");

                        // Validar dados no SAP
                        atualizador(50, "Validando dados no SAP...");
                        var validacao = _processamentoController.ValidarDadosSAP(linhasSelecionadas);

                        atualizador(80, "Processando resultados...");

                        if (validacao.Valida)
                        {
                            Application.SBO_Application.MessageBox(
                                $"✓ Validação concluída com sucesso!\n\n" +
                                $"{linhasSelecionadas.Count} linhas validadas e prontas para processamento.",
                                1, "Ok", "", "");

                            MarcarLinhasValidadas(linhasSelecionadas);
                        }
                        else
                        {
                            string erros = string.Join("\n", validacao.Erros.Take(10));
                            if (validacao.Erros.Count > 10)
                                erros += $"\n... e mais {validacao.Erros.Count - 10} erros";

                            string avisos = validacao.Avisos.Any() ?
                                $"\n\nAvisos:\n{string.Join("\n", validacao.Avisos.Take(5))}" : "";

                            int resposta = Application.SBO_Application.MessageBox(
                                $"Foram encontrados problemas na validação:\n\n{erros}{avisos}\n\n" +
                                $"Deseja continuar mesmo assim?",
                                2, "Sim", "Não", "");

                            if (resposta == 1)
                            {
                                MarcarLinhasValidadas(linhasSelecionadas);
                            }
                        }

                        atualizador(100, "Validação concluída!");
                    });
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro na validação: {ex.Message}", 1, "Ok", "", "");
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void BtnProcessar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                var linhasSelecionadas = ObterLinhasSelecionadasOtimizado();

                if (linhasSelecionadas.Count == 0)
                {
                    Application.SBO_Application.MessageBox("Selecione pelo menos uma linha para processar.", 1, "Ok", "", "");
                    return;
                }

                int resposta = Application.SBO_Application.MessageBox(
                    $"Confirma o processamento de {linhasSelecionadas.Count} documento(s)?\n\n" +
                    $"Valor total: {_valorTotalSelecionado:C}\n" +
                    $"Data de Lançamento: {TxtDtLanc.Value}\n" +
                    $"Data do Documento: {TxtDtDoc.Value}",
                    2, "Sim", "Não", "");

                if (resposta == 1)
                {
                    AbrirFormularioResultado(linhasSelecionadas);
                }
            }
            catch (Exception ex)
            {
               // Application.SBO_Application.MessageBox($"Erro: {ex.Message}", 1, "Ok", "", "");
            }
        }

        #endregion

        #region Navegação

        private void BtnVoltar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (_totalLinhasSelecionadas > 0)
                {
                    int resposta = Application.SBO_Application.MessageBox(
                        "Existem linhas selecionadas. Deseja voltar sem processar?",
                        2, "Sim", "Não", "");

                    if (resposta != 1)
                    {
                        BubbleEvent = false;
                        return;
                    }
                }

                if (!string.IsNullOrEmpty(_formOrigemUID))
                {
                    FormManager.TrazerParaFrente(_formOrigemUID);
                }
                else
                {
                    //FormGrupoLote formGrupo = new FormGrupoLote(UIAPIRawForm.UniqueID);
                    //formGrupo.SetGrupoCode(_grupoCode);
                    //formGrupo.Show();
                }

                UIAPIRawForm.Close();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao voltar: {ex.Message}", BoMessageTime.bmt_Short, true);
            }
        }

        private void BtnCancelar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (_totalLinhasSelecionadas > 0)
                {
                    int resposta = Application.SBO_Application.MessageBox(
                        "Deseja realmente cancelar o processamento?",
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
                Application.SBO_Application.SetStatusBarMessage($"Erro: {ex.Message}", BoMessageTime.bmt_Short, true);
            }
        }

        private void AbrirFormularioResultado(List<LinhaImportacao> linhasSelecionadas)
        {
            try
            {
                // ✅ MUDANÇA 1: Usar FormProcessamentoNFS ao invés de FormResultadoProcessamento
                FormProcessamentoNFS formProcessamento = new FormProcessamentoNFS();

                // ✅ MUDANÇA 2: Status deve ser ProcessandoOuEmAndamento, não Novo
                _grupo.Status = StatusGrupo.EmProcessamento; // ou ProcessandoOuEmAndamento se existir

                // ✅ MUDANÇA 3: Atualizar linhas selecionadas no grupo
                // Se linhasSelecionadas for fornecida, usar apenas elas
                // Se não, usar todas as linhas do grupo
                if (linhasSelecionadas != null && linhasSelecionadas.Count > 0)
                {
                    _grupo.Linhas = linhasSelecionadas;
                }
                // Se linhasSelecionadas for null/vazia, manter _grupo.Linhas como está
                formProcessamento.SetGrupo(_grupo);

                formProcessamento.Show();

                // ✅ MUDANÇA 4: Passar grupo completo

                // ✅ MUDANÇA 5: Registrar no FormManager para controle
                FormManager.RegistrarFormulario(
                    formProcessamento.UIAPIRawForm.UniqueID,
                    "Processamento NFS-e",
                    false,
                    UIAPIRawForm.UniqueID,
                    _grupo.Code,
                    3 // Etapa 3 = Processamento
                );

           

                UIAPIRawForm.Close();
            }
            catch (Exception ex)
            {
                //throw new Exception($"Erro ao abrir formulário de processamento: {ex.Message}", ex);
            }
        }

        #endregion

        #region Métodos Auxiliares Otimizados

        private List<LinhaImportacao> ObterLinhasSelecionadas()
        {
            return ObterLinhasSelecionadasOtimizado();
        }

        private List<LinhaImportacao> ObterLinhasSelecionadasOtimizado()
        {
            if (_linhasSelecionadasIndex != null && _linhasSelecionadasIndex.Count > 0)
            {
                return _linhasSelecionadasIndex
                    .Where(i => _linhasIndexadas.ContainsKey(i))
                    .Select(i => _linhasIndexadas[i])
                    .ToList();
            }

            // Fallback para método tradicional
            var linhasSelecionadas = new List<LinhaImportacao>();
            var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");

            for (int i = 0; i < dt.Rows.Count && i < _linhas.Count; i++)
            {
                if (dt.GetValue("Sel", i).ToString() == "Y")
                {
                    var linha = _linhas[i];
                    linha.Selecionada = true;
                    linhasSelecionadas.Add(linha);
                }
            }

            return linhasSelecionadas;
        }

        private void MarcarLinhasValidadas(List<LinhaImportacao> linhas)
        {
            try
            {
                UIAPIRawForm.Freeze(true);

                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");
                var codesValidados = new HashSet<string>(linhas.Select(l => l.Code));

                // Usar busca otimizada com HashSet
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string code = dt.GetValue("Code", i).ToString();
                    if (codesValidados.Contains(code))
                    {
                        dt.SetValue("Status", i, "✓ Validado");
                    }
                }

                MtxLinhas.LoadFromDataSource();
                ColorirStatusMatrix();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao marcar validação: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void ColorirStatusMatrix()
        {
            var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");

            // Colorir apenas linhas visíveis para melhor performance
            int visibleRows = Math.Min(MtxLinhas.RowCount, 50);

            for (int i = 1; i <= visibleRows; i++)
            {
                try
                {
                    string status = dt.GetValue("Status", i - 1).ToString();
                    var statusCell = MtxLinhas.Columns.Item("Status").Cells.Item(i).Specific as SAPbouiCOM.EditText;

                    if (statusCell != null)
                    {
                        if (status.Contains("✓"))
                            statusCell.ForeColor = 2263842; // Verde escuro
                        else if (status.Contains("✗") || status.Contains("Erro"))
                            statusCell.ForeColor = 255; // Vermelho
                        else
                            statusCell.ForeColor = -1; // Cor padrão
                    }
                }
                catch
                {
                    continue;
                }
            }
        }

        private void MostrarProgressoCarregamento(bool mostrar)
        {
            try
            {
                var lblItem = UIAPIRawForm.Items.Item("lblStImp");
                var lblInstrucaoItem = UIAPIRawForm.Items.Item("lblInst");

                lblItem.Visible = mostrar;
                lblInstrucaoItem.Visible = !mostrar;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao atualizar progresso: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void AtualizarInterface()
        {
            try
            {

                bool temLinhas = _linhas != null && _linhas.Count > 0;

                CmbFiltro.Item.Enabled = temLinhas;
                BtnValidar.Item.Enabled = temLinhas && _totalLinhasSelecionadas > 0;
                BtnProcessar.Item.Enabled = temLinhas && _totalLinhasSelecionadas > 0;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao atualizar interface: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private string ObterDescricaoStatus(StatusLinha status)
        {
            switch (status)
            {
                case StatusLinha.Pendente: return "Pendente";
                case StatusLinha.Processando: return "Processando";
                case StatusLinha.Sucesso: return "✓ Sucesso";
                case StatusLinha.Erro: return "✗ Erro";
                case StatusLinha.Ignorada: return "Ignorada";
                default: return "Desconhecido";
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

        #endregion

        #region Eventos do Formulário

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {
            try
            {
                UIAPIRawForm.Title = "Processar NFS-e em Lote - Etapa 2: Seleção de Linhas";

                if (!FormManager.FormularioEstaAberto(UIAPIRawForm.UniqueID))
                {
                    FormManager.RegistrarFormulario(
                        UIAPIRawForm.UniqueID,
                        "Visualização de Linhas - Etapa 2",
                        isMaster: false,
                        masterFormUID: _formOrigemUID ?? FormManager.FORM_SELECIONAR_GRUPO,
                        grupoCode: _grupoCode,
                        etapa: 2
                    );
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
            _formFechando = true;

            try
            {
                // Cancelar operações em andamento
                _cancellationTokenSource?.Cancel();

                // Parar timer UI
                PararTimerUI();

                // Limpar DataTable
                if (DataTableExists("dtLinhas"))
                {
                    var dt = UIAPIRawForm.DataSources.DataTables.Item("dtLinhas");
                    dt.Rows.Clear();
                }

                // Limpar coleções
                _linhasIndexadas?.Clear();
                _linhasSelecionadasIndex?.Clear();
                _linhas?.Clear();

                // Fechar ProgressBar se estiver aberta
                ProgressBarHelper.Instance.Fechar();

                // Remover do FormManager
                FormManager.RemoverFormulario(UIAPIRawForm.UniqueID);

                // Dispose do CancellationTokenSource
                _cancellationTokenSource?.Dispose();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao fechar: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            try
            {
                if (MtxLinhas != null)
                {
                    MtxLinhas.AutoResizeColumns();
                }
            }
            catch { }
        }

        #endregion

    }
}