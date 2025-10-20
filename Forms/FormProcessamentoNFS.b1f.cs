using ItTech.Tool.AddonNFS.Controllers;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Services;
using ItTech.Tool.AddonNFS.Utils;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Application = SAPbouiCOM.Framework.Application;

namespace ItTech.Tool.AddonNFS.Forms
{
    [FormAttribute("ITTECH_NFS_PRO", "Forms/FormProcessamentoNFS.b1f")]
    class FormProcessamentoNFS : UserFormBase
    {
        #region Campos

        private GrupoLote _grupo;
        private ProcessamentoNFSController _processamentoController;
        private ServiceLayerInvoiceClient _serviceLayerClient;
        private bool _processamentoEmAndamento = false;
        private bool _configurado = false;
        private CancellationTokenSource _cancellationTokenSource;
        private readonly object _lockTotais = new object();
        private string _formOriginId;

        // Controles
        private StaticText lblTitulo, lblStatus;
        private EditText txtStatus, txtGrupo, txtDtLanc, txtDtDoc;
        private ComboBox cboFiltro;
        private Button btnProc, btnExpErr, btnVoltar, btnFechar;
        private Matrix oGrid;

        // Estatísticas
        private int _totalSelecionadas = 0;
        private decimal _valorTotalSelecionado = 0;
        
        // Cache para otimização
        private Dictionary<string, StatusLinha> _ultimoStatusCache = new Dictionary<string, StatusLinha>();
        private bool _zebraStripeAplicado = false;

        // Cores de status - SUPER SIMPLES E CLARAS
        private readonly Dictionary<StatusLinha, int> CORES_STATUS = new Dictionary<StatusLinha, int>
        {
            { StatusLinha.Sucesso, 13172680 },   // Verde bem claro (200,255,200)
            { StatusLinha.Erro, 13158655 },      // Rosa bem claro (255,200,200)
            { StatusLinha.Pendente, 16763080 }   // Azul bem claro (200,200,255)
        };

        #endregion

        #region Inicialização

        public FormProcessamentoNFS(string formOriginId = null)
        {
            _formOriginId = formOriginId;
            ConfigurarSupressaoExcecoes();
        }

        private void ConfigurarSupressaoExcecoes()
        {
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                System.Diagnostics.Debug.WriteLine($"Exceção suprimida: {(e.ExceptionObject as Exception)?.Message}");
        }

        public override void OnInitializeComponent()
        {
            // Inicializar controles
            lblTitulo = GetControl<StaticText>("lblTitulo");
            lblStatus = GetControl<StaticText>("lblStatus");
            txtStatus = GetControl<EditText>("txtStatus");
            txtGrupo = GetControl<EditText>("txtGrupo");
            txtDtLanc = GetControl<EditText>("txtDtLanc");
            txtDtDoc = GetControl<EditText>("txtDtDoc");
            cboFiltro = GetControl<ComboBox>("cboFiltro");
            btnProc = GetControl<Button>("btnProc");
            btnExpErr = GetControl<Button>("btnExpErr");
            btnVoltar = GetControl<Button>("btnVoltar");
            btnFechar = GetControl<Button>("btnFechar");
            oGrid = GetControl<Matrix>("oGrid");

            // Configurar eventos principais
            cboFiltro.ComboSelectAfter += CboFiltro_ComboSelectAfter;
            oGrid.ClickAfter += Matrix_ClickAfter;
            btnProc.ClickBefore += BtnProcessar_ClickBefore;
            btnExpErr.ClickBefore += BtnExportarErros_ClickBefore;
            btnVoltar.ClickBefore += BtnVoltar_ClickBefore;
            btnFechar.ClickBefore += BtnFechar_ClickBefore;

            OnCustomInitialize();
        }

        private T GetControl<T>(string itemId) where T : class
        {
            return GetItem(itemId).Specific as T;
        }

        private void OnCustomInitialize()
        {
            try
            {
                if (_configurado) return;

                var company = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                _serviceLayerClient = new ServiceLayerInvoiceClient();
                _processamentoController = new ProcessamentoNFSController(company, _serviceLayerClient);

                ConfigurarFormulario();
                _configurado = true;
            }
            catch (Exception ex)
            {
                MostrarMensagem($"Erro ao inicializar: {ex.Message}");
            }
        }

        private void ConfigurarFormulario()
        {
            ConfigurarDataTable();
            ConfigurarMatrix();
            ConfigurarFiltros();
        }

        #endregion

        #region Configuração

        private void ConfigurarDataTable()
        {
            try
            {
                var dt = UIAPIRawForm.DataSources.DataTables.Add("dtResult");

                // Adicionar colunas
                dt.Columns.Add("Linha", BoFieldsType.ft_Integer);
                dt.Columns.Add("Proc", BoFieldsType.ft_AlphaNumeric, 1);
                dt.Columns.Add("Filial", BoFieldsType.ft_Integer);
                dt.Columns.Add("CodCli", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("Cliente", BoFieldsType.ft_AlphaNumeric, 200);
                dt.Columns.Add("CodItem", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("Descricao", BoFieldsType.ft_AlphaNumeric, 200);
                dt.Columns.Add("Utiliz", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("DocEntry", BoFieldsType.ft_Integer);
                dt.Columns.Add("DocNum", BoFieldsType.ft_Integer);
                dt.Columns.Add("Status", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("CodImp", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("CodSeq", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("Condicao", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("ObsNF", BoFieldsType.ft_AlphaNumeric, 254);
                dt.Columns.Add("TipoTrib", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("Valor", BoFieldsType.ft_Float);
                dt.Columns.Add("Mensagem", BoFieldsType.ft_AlphaNumeric, 254);
                dt.Columns.Add("Code", BoFieldsType.ft_AlphaNumeric, 50);
            }
            catch { /* DataTable já existe */ }
        }

        private void ConfigurarMatrix()
        {
            try
            {
                // Vincular colunas
                foreach (Column col in oGrid.Columns)
                {
                    if (col.UniqueID != "#")
                        col.DataBind.Bind("dtResult", col.UniqueID);
                }
                oGrid.Columns.Item("#").DataBind.Bind("dtResult", "Linha");

                // LinkedButtons
                ConfigurarLinkedButton("CodCli", BoLinkedObject.lf_BusinessPartner);
                ConfigurarLinkedButton("CodItem", BoLinkedObject.lf_Items);
                ConfigurarLinkedButton("DocEntry", BoLinkedObject.lf_Invoice);
            }
            catch (Exception ex)
            {
                MostrarMensagem($"Erro ao configurar matrix: {ex.Message}");
            }
        }

        private void ConfigurarLinkedButton(string coluna, BoLinkedObject tipo)
        {
            var link = (LinkedButton)oGrid.Columns.Item(coluna).ExtendedObject;
            link.LinkedObject = tipo;
        }

        private void ConfigurarFiltros()
        {
            try
            {
                // Limpar valores existentes
                if (cboFiltro.ValidValues.Count > 0)
                {
                    for (int i = cboFiltro.ValidValues.Count - 1; i >= 0; i--)
                    {
                        cboFiltro.ValidValues.Remove(i, BoSearchKey.psk_Index);
                    }
                }

                // Adicionar novos valores
                cboFiltro.ValidValues.Add("TODOS", "Mostrar Todos");
                cboFiltro.ValidValues.Add("SUCESSO", "Com Sucesso");
                cboFiltro.ValidValues.Add("ERRO", "Com Erro");
                cboFiltro.ValidValues.Add("PENDENTE", "Pendentes");

                cboFiltro.Select("TODOS", BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao configurar filtros: {ex.Message}");
            }
        }

        #endregion

        #region Entrada de Dados

        public void SetGrupo(GrupoLote grupo)
        {
            if (grupo == null)
            {
                MostrarMensagem("Grupo não pode ser nulo.");
                return;
            }

            _grupo = grupo;

            try
            {
                CarregarDados();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao carregar dados: {ex.Message}");
                MostrarErro($"Erro ao carregar dados: {ex.Message}");
            }
        }

        private void CarregarDados()
        {
            ExecutarComFreeze(() =>
            {
                // Cabeçalho
                txtGrupo.Value = _grupo.Nome;
                txtDtLanc.Value = _grupo.DataLancamento.ToString("dd/MM/yyyy");
                txtDtDoc.Value = _grupo.DataDocumento.ToString("dd/MM/yyyy");

                // Carregar linhas se necessário
                if (_grupo.Linhas == null || _grupo.Linhas.Count == 0)
                {
                    var controller = new GrupoLoteController(
                        (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany()
                    );
                    _grupo.Linhas = controller.ObterLinhasGrupo(_grupo.Code);
                }

                CarregarMatrix();
            });
        }

        private void CarregarMatrix()
        {
            try
            {
                // Mostrar indicador de carregamento
                Application.SBO_Application.StatusBar.SetText("Carregando dados... Por favor aguarde.", 
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                
                var dt = GetDataTable();
                if (dt == null) return;

                dt.Rows.Clear();

                if (_grupo.Linhas == null || _grupo.Linhas.Count == 0)
                {
                    oGrid.Clear();
                    Application.SBO_Application.StatusBar.SetText("Nenhum dado para exibir.", 
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    return;
                }

                var linhasOrdenadas = OrdenarLinhas(_grupo.Linhas);
                
                // Adicionar todas as linhas de uma vez para melhor performance
                int totalLinhas = linhasOrdenadas.Count;
                bool mostrarProgresso = totalLinhas > 100;
                
                if (mostrarProgresso)
                {
                    Application.SBO_Application.StatusBar.SetText($"Preparando {totalLinhas} registros...", 
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }

                // Adicionar linhas em lotes para melhor performance
                const int BATCH_SIZE = 50;
                for (int batch = 0; batch < totalLinhas; batch += BATCH_SIZE)
                {
                    var lote = linhasOrdenadas.Skip(batch).Take(BATCH_SIZE);
                    foreach (var linha in lote)
                    {
                        AdicionarLinhaDataTable(dt, linha);
                    }
                    
                    // Atualizar progresso apenas em lotes grandes
                    if (mostrarProgresso && batch % 100 == 0)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            $"Carregando... {batch}/{totalLinhas} registros", 
                            BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    }
                }

                // Carregar dados no grid de uma vez
                oGrid.Clear();
                oGrid.LoadFromDataSource();
                
                // Aplicar estilos apenas se necessário
                if (totalLinhas > 0)
                {
                    AplicarEstiloMatrixOtimizado();
                }
                
                // Recalcular totais após carregar os dados
                RecalcularTotais();
                AtualizarStatus();
                AtualizarBotoes();
                
                // Indicar conclusão do carregamento
                Application.SBO_Application.StatusBar.SetText($"Dados carregados com sucesso! {totalLinhas} registros.", 
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro em CarregarMatrix: {ex.Message}");
                Application.SBO_Application.StatusBar.SetText("Erro ao carregar dados.", 
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private List<LinhaImportacao> OrdenarLinhas(List<LinhaImportacao> linhas)
        {
            if (linhas == null || linhas.Count == 0)
                return new List<LinhaImportacao>();

            return linhas.OrderBy(l => ObterOrdemStatus(l.Status))
                        .ThenBy(l => l.NumeroLinha)
                        .ToList();
        }

        private int ObterOrdemStatus(StatusLinha status)
        {
            switch (status)
            {
                case StatusLinha.Pendente: return 1;
                case StatusLinha.Erro: return 2;
                case StatusLinha.Sucesso: return 3;
                default: return 4;
            }
        }

        private void AdicionarLinhaDataTable(DataTable dt, LinhaImportacao linha)
        {
            dt.Rows.Add();
            int i = dt.Rows.Count - 1;

            dt.SetValue("Linha", i, linha.NumeroLinha);
            // Por padrão, marcar como "Y" apenas se estiver pendente
            dt.SetValue("Proc", i, linha.Status == StatusLinha.Pendente ? "Y" : "N");
            dt.SetValue("Filial", i, linha.Filial);
            dt.SetValue("CodCli", i, linha.CodigoCliente);
            dt.SetValue("Cliente", i, Truncar(linha.NomeCliente, 200));
            dt.SetValue("CodItem", i, linha.CodigoItem);
            dt.SetValue("Descricao", i, Truncar(linha.DescricaoItem, 200));
            dt.SetValue("Utiliz", i, linha.Utilizacao ?? "");
            dt.SetValue("DocEntry", i, linha.DocEntry ?? 0);
            dt.SetValue("DocNum", i, linha.DocNum ?? 0);
            dt.SetValue("Status", i, ObterTextoStatus(linha.Status));
            dt.SetValue("CodImp", i, linha.CodigoImposto ?? "");
            dt.SetValue("CodSeq", i, linha.CodSeq ?? "");
            dt.SetValue("Condicao", i, linha.CondicaoPagamento ?? "");
            dt.SetValue("ObsNF", i, Truncar(linha.ObservacaoNF, 254));
            dt.SetValue("TipoTrib", i, linha.TipoTributacao ?? "");
            dt.SetValue("Valor", i, Convert.ToDouble(linha.Valor));
            dt.SetValue("Mensagem", i, Truncar(linha.MensagemErro, 254));
            dt.SetValue("Code", i, linha.Code);
        }

        private string ObterTextoStatus(StatusLinha status)
        {
            switch (status)
            {
                case StatusLinha.Sucesso: return "✓ Sucesso";
                case StatusLinha.Erro: return "✗ Erro";
                default: return "○ Pendente";
            }
        }

        #endregion

        #region Processamento

        private void IniciarProcessamento()
        {
            try
            {
                if (_grupo == null || _grupo.Linhas == null)
                {
                    MostrarMensagem("Não há dados para processar.");
                    return;
                }

                _processamentoEmAndamento = true;
                
                // Cancelar operação anterior se existir
                _cancellationTokenSource?.Cancel();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = new CancellationTokenSource();
                
                AtualizarBotoes();

                var linhasParaProcessar = ObterLinhasSelecionadas();

                if (linhasParaProcessar.Count == 0)
                {
                    MostrarMensagem("Não há documentos selecionados para processar.");
                    _processamentoEmAndamento = false;
                    AtualizarInterface();
                    return;
                }

                // Usar await para garantir tratamento de erros
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await ProcessarLinhasAsync(linhasParaProcessar);
                    }
                    catch (OperationCanceledException)
                    {
                        MostrarStatus("Processamento cancelado");
                    }
                    catch (Exception ex)
                    {
                        MostrarErro($"Erro no processamento: {ex.Message}");
                    }
                    finally
                    {
                        _processamentoEmAndamento = false;
                        AtualizarInterface();
                    }
                }, _cancellationTokenSource.Token);
            }
            catch (Exception ex)
            {
                MostrarErro($"Erro ao iniciar processamento: {ex.Message}");
                _processamentoEmAndamento = false;
                AtualizarInterface();
            }
        }

        private async Task ProcessarLinhasAsync(List<LinhaImportacao> linhas)
        {
            if (_grupo == null) return;

            var progress = ProgressBarHelper.Instance;

            try
            {
                // Configuração
                progress.Criar($"Processando {linhas.Count} documentos...", 100);

                // Validação
                progress.Atualizar(10, "Validando dados...");
                var validacao = _processamentoController.ValidarDadosSAP(linhas);

                if (!validacao.Valida)
                {
                    progress.Fechar();
                    MostrarMensagem($"Erros encontrados:\n{string.Join("\n", validacao.Erros.Take(5))}");
                    return;
                }

                // Conexão
                progress.Atualizar(20, "Conectando ao Service Layer...");
                if (!await _serviceLayerClient.ConnectAsync(_cancellationTokenSource.Token))
                {
                    throw new Exception("Falha ao conectar com Service Layer");
                }

                // Processamento em lotes
                int total = linhas.Count;
                int processadas = 0;
                int sucessos = 0;
                int erros = 0;

                const int LOTE_SIZE = 25;
                const int LIMITE_TEMPO_REAL = 50;
                var resultadosAcumulados = new List<ResultadoProcessamento>();
                bool atualizarEmTempoReal = true;

                for (int i = 0; i < total; i += LOTE_SIZE)
                {
                    var lote = linhas.Skip(i).Take(LOTE_SIZE).ToList();
                    int progresso = 20 + (int)((processadas / (double)total) * 70);

                    progress.Atualizar(progresso, $"Processando {processadas + 1} a {i + lote.Count} de {total}...");

                    
                    var resultados = await _processamentoController.ProcessarLinhasAsync(_grupo, lote);

                    // Lógica de atualização inteligente
                    if (processadas < LIMITE_TEMPO_REAL)
                    {
                        // Primeiros 50: atualiza em tempo real
                        AtualizarResultados(resultados);
                        
                        // Verifica se passou do limite neste lote
                        if (processadas + lote.Count >= LIMITE_TEMPO_REAL && atualizarEmTempoReal)
                        {
                            atualizarEmTempoReal = false;
                            Application.SBO_Application.StatusBar.SetText(
                                "Processando lote maior... Interface será atualizada ao final.", 
                                BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                        }
                    }
                    else
                    {
                        // Após 50: apenas acumula
                        resultadosAcumulados.AddRange(resultados);
                    }

                    sucessos += resultados.Count(r => r.Sucesso);
                    erros += resultados.Count(r => !r.Sucesso);
                    processadas += lote.Count;

                    // Verificar cancelamento
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        throw new OperationCanceledException();
                    }
                    
                    // Pequeno delay apenas se necessário para não sobrecarregar
                    if (lote.Count == LOTE_SIZE)
                        await Task.Delay(50, _cancellationTokenSource.Token);
                }

                // Se acumulou resultados, atualiza tudo de uma vez
                if (resultadosAcumulados.Count > 0)
                {
                    progress.Atualizar(95, "Atualizando interface com todos os resultados...");
                    Application.SBO_Application.StatusBar.SetText(
                        $"Atualizando {resultadosAcumulados.Count} resultados... Por favor aguarde.", 
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    
                    AtualizarResultados(resultadosAcumulados);
                }

                // Finalização
                progress.Atualizar(100, "Concluído!");
                progress.Fechar();

                MostrarMensagem($"Processamento concluído!\n\n✅ Sucessos: {sucessos}\n❌ Erros: {erros}");
            }
            catch (Exception ex)
            {
                progress.Fechar();
                MostrarErro($"Erro no processamento: {ex.Message}");
            }
            finally
            {
                _processamentoEmAndamento = false;
                // Apenas atualizar botões, não recarregar tudo
                AtualizarBotoes();
            }
        }

        private void AtualizarResultados(List<ResultadoProcessamento> resultados)
        {
            try
            {
                if (_grupo == null || _grupo.Linhas == null) return;

                var dt = GetDataTable();
                if (dt == null || resultados == null || resultados.Count == 0) return;

                // Mostrar indicador apenas se houver muitos resultados
                bool mostrarProgresso = resultados.Count > 50;
                if (mostrarProgresso)
                {
                    Application.SBO_Application.StatusBar.SetText($"Atualizando {resultados.Count} resultados... Por favor aguarde.", 
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }

                // Criar índice para lookup O(1) ao invés de O(n²)
                var codeToIndex = new Dictionary<string, int>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var code = dt.GetValue("Code", i)?.ToString();
                    if (!string.IsNullOrEmpty(code))
                    {
                        codeToIndex[code] = i;
                    }
                }

                // Criar índice para linhas também
                var linhasPorCode = _grupo.Linhas.ToDictionary(l => l.Code, l => l);

                // Atualizar DataTable com resultados - O(n) ao invés de O(n²)
                foreach (var resultado in resultados)
                {
                    if (codeToIndex.TryGetValue(resultado.CodigoLinha, out int index))
                    {
                        dt.SetValue("Status", index, ObterTextoStatus(resultado.Sucesso ? StatusLinha.Sucesso : StatusLinha.Erro));
                        dt.SetValue("Mensagem", index, Truncar(resultado.Mensagem, 254));
                        dt.SetValue("Proc", index, "N");

                        if (resultado.DocEntry.HasValue)
                        {
                            dt.SetValue("DocEntry", index, resultado.DocEntry.Value);
                            dt.SetValue("DocNum", index, resultado.DocNum ?? resultado.DocEntry.Value);
                        }

                        // Atualizar objeto - também O(1)
                        if (linhasPorCode.TryGetValue(resultado.CodigoLinha, out var linha))
                        {
                            linha.Status = resultado.Sucesso ? StatusLinha.Sucesso : StatusLinha.Erro;
                            linha.MensagemErro = resultado.Mensagem;
                            linha.DocEntry = resultado.DocEntry;
                            linha.DocNum = resultado.DocNum;
                        }
                    }
                }

                // Atualizar grid apenas uma vez
                oGrid.LoadFromDataSource();
                
                // Aplicar estilo apenas nas linhas que mudaram
                AplicarEstiloMatrixOtimizado(resultados);
                
                // Atualizar totais e status apenas
                RecalcularTotais();
                AtualizarStatus();
                AtualizarBotoes();
                
                // Indicar conclusão se estava mostrando progresso
                if (mostrarProgresso)
                {
                    Application.SBO_Application.StatusBar.SetText("Resultados atualizados com sucesso!", 
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro em AtualizarResultados: {ex.Message}");
                Application.SBO_Application.StatusBar.SetText("Erro ao atualizar resultados.", 
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Interface

        private void AtualizarInterface()
        {
            RecalcularTotais();
            AtualizarStatus();
            AtualizarBotoes();
        }

        private void RecalcularTotais()
        {
            lock (_lockTotais)
            {
                try
                {
                    _totalSelecionadas = 0;
                    _valorTotalSelecionado = 0;

                    if (_grupo == null || _grupo.Linhas == null) return;

                    var dt = GetDataTable();
                    if (dt == null || dt.Rows.Count == 0) return;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string procValue = dt.GetValue("Proc", i)?.ToString() ?? "N";

                        if (procValue == "Y")
                        {
                            var code = dt.GetValue("Code", i)?.ToString();
                            if (!string.IsNullOrEmpty(code))
                            {
                                var linha = _grupo.Linhas.FirstOrDefault(l => l.Code == code);

                                if (linha != null && linha.Status == StatusLinha.Pendente)
                                {
                                    _totalSelecionadas++;
                                    _valorTotalSelecionado += linha.Valor;
                                }
                            }
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"Totais recalculados: {_totalSelecionadas} selecionadas, valor: {_valorTotalSelecionado:C}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erro em RecalcularTotais: {ex.Message}");
                }
            }
        }

        private void AtualizarStatus()
        {
            try
            {
                if (txtStatus == null || txtStatus.Item == null || _grupo == null) return;

                var stats = ObterEstatisticas();
                string status = $"Status: {stats.Item1} documentos | ✓ {stats.Item2} | ✕ {stats.Item3} | ⚠ {stats.Item4}";

                if (_totalSelecionadas > 0)
                {
                    status += $" | 📌 {_totalSelecionadas} selecionadas ({_valorTotalSelecionado:C})";
                }

                // Forçar atualização do campo
                txtStatus.Item.Update();
                txtStatus.Value = status;

                System.Diagnostics.Debug.WriteLine($"Status atualizado: {status}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro em AtualizarStatus: {ex.Message}");
            }
        }

        private void AtualizarBotoes()
        {
            try
            {
                if (btnProc == null || btnProc.Item == null) return;

                // Habilitar botão processar apenas se há seleções e não está processando
                btnProc.Item.Enabled = _totalSelecionadas > 0 && !_processamentoEmAndamento;

                // Verificar se há erros
                bool temErros = false;
                if (_grupo != null && _grupo.Linhas != null)
                {
                    temErros = _grupo.Linhas.Any(l => l.Status == StatusLinha.Erro);
                }

                if (btnExpErr != null && btnExpErr.Item != null)
                {
                    btnExpErr.Item.Enabled = temErros;
                }

                System.Diagnostics.Debug.WriteLine($"Botões atualizados - Processar: {btnProc.Item.Enabled}, Exportar: {temErros}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro em AtualizarBotoes: {ex.Message}");
            }
        }

        private Tuple<int, int, int, int> ObterEstatisticas()
        {
            if (_grupo == null || _grupo.Linhas == null)
                return Tuple.Create(0, 0, 0, 0);

            int total = _grupo.Linhas.Count;
            int sucessos = _grupo.Linhas.Count(l => l.Status == StatusLinha.Sucesso);
            int erros = _grupo.Linhas.Count(l => l.Status == StatusLinha.Erro);
            int pendentes = _grupo.Linhas.Count(l => l.Status == StatusLinha.Pendente);

            return Tuple.Create(total, sucessos, erros, pendentes);
        }

        private void AplicarEstiloMatrixOtimizado(List<ResultadoProcessamento> resultadosRecentes = null)
        {
            try
            {
                if (oGrid == null || oGrid.RowCount == 0) return;
                if (_grupo == null || _grupo.Linhas == null) return;

                var dt = GetDataTable();
                if (dt == null) return;

                // Mostrar indicador apenas se houver muitas linhas
                bool mostrarProgresso = oGrid.RowCount > 100;
                if (mostrarProgresso)
                {
                    Application.SBO_Application.StatusBar.SetText("Aplicando formatação visual...", 
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }

                // Criar dicionário para lookup O(1)
                var linhasPorCode = _grupo.Linhas.ToDictionary(l => l.Code, l => l);
                
                // Se temos resultados recentes, processar apenas essas linhas
                if (resultadosRecentes != null && resultadosRecentes.Count > 0)
                {
                    var codesAlterados = new HashSet<string>(resultadosRecentes.Select(r => r.CodigoLinha));
                    
                    for (int i = 1; i <= oGrid.RowCount; i++)
                    {
                        int dtIndex = i - 1;
                        if (dtIndex >= dt.Rows.Count) continue;

                        var code = dt.GetValue("Code", dtIndex)?.ToString();
                        if (!string.IsNullOrEmpty(code) && codesAlterados.Contains(code))
                        {
                            if (linhasPorCode.TryGetValue(code, out var linha))
                            {
                                AplicarCorLinha(i, linha);
                                _ultimoStatusCache[code] = linha.Status;
                            }
                        }
                    }
                }
                else
                {
                    // Primeira vez ou recarregamento completo
                    _ultimoStatusCache.Clear();
                    
                    for (int i = 1; i <= oGrid.RowCount; i++)
                    {
                        int dtIndex = i - 1;
                        if (dtIndex >= dt.Rows.Count) continue;

                        var code = dt.GetValue("Code", dtIndex)?.ToString();
                        if (!string.IsNullOrEmpty(code) && linhasPorCode.TryGetValue(code, out var linha))
                        {
                            AplicarCorLinha(i, linha);
                            _ultimoStatusCache[code] = linha.Status;
                        }
                    }
                    
                    // Aplicar zebra stripe apenas uma vez
                    if (!_zebraStripeAplicado)
                    {
                        AplicarZebraStripe();
                        _zebraStripeAplicado = true;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro em AplicarEstiloMatrixOtimizado: {ex.Message}");
            }
        }
        
        private void AplicarCorLinha(int row, LinhaImportacao linha)
        {
            try
            {
                // Checkbox - desabilitar visualmente se não pendente
                if (linha.Status != StatusLinha.Pendente)
                {
                    oGrid.CommonSetting.SetCellBackColor(row, 1, Color.LightGray.ToArgb());
                }
                
                // Cor de status
                if (CORES_STATUS.TryGetValue(linha.Status, out int cor))
                {
                    oGrid.CommonSetting.SetCellBackColor(row, 10, cor);
                }
            }
            catch { }
        }
        
        private void AplicarZebraStripe()
        {
            try
            {
                int corBranca = Color.White.ToArgb();
                int corCinza = Color.FromArgb(245, 247, 250).ToArgb();
                
                for (int i = 1; i <= oGrid.RowCount; i++)
                {
                    oGrid.CommonSetting.SetRowBackColor(i, (i - 1) % 2 == 0 ? corBranca : corCinza);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao aplicar zebra stripe: {ex.Message}");
            }
        }

        #endregion

        #region Eventos

        private void Matrix_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            ExecutarComFreeze(() =>
            {
                try
                {
                    // Selecionar a linha clicada
                    if (pVal.Row > 0 && pVal.Row <= oGrid.RowCount)
                    {
                        oGrid.SelectRow(pVal.Row, true, false);
                    }

                    // Se clicou na coluna Proc, fazer o toggle
                    if (pVal.ColUID == "Proc" && pVal.Row > 0)
                    {
                        if (_grupo == null || _grupo.Linhas == null) return;

                        var dt = GetDataTable();
                        if (dt == null) return;

                        int index = pVal.Row - 1;
                        if (index < 0 || index >= dt.Rows.Count) return;

                        var code = dt.GetValue("Code", index).ToString();
                        var linha = _grupo.Linhas.FirstOrDefault(l => l.Code == code);

                        // Se a linha está pendente, pode alternar
                        if (linha != null && linha.Status == StatusLinha.Pendente)
                        {
                            // Toggle do valor
                            string valorAtual = dt.GetValue("Proc", index).ToString();
                            string novoValor = valorAtual == "Y" ? "N" : "Y";
                            dt.SetValue("Proc", index, novoValor);

                            // Atualizar visualmente a grid
                            oGrid.LoadFromDataSource();

                            // Recalcular totais e atualizar interface
                            RecalcularTotais();
                            AtualizarStatus();
                            AtualizarBotoes();
                        }
                        else if (linha != null)
                        {
                            // Se não está pendente, garantir que fica como "N"
                            dt.SetValue("Proc", index, "N");
                            oGrid.LoadFromDataSource();
                            MostrarStatus("⚠️ Documentos já processados não podem ser alterados");
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erro em Matrix_ClickAfter: {ex.Message}");
                }
            });
        }

        private void CboFiltro_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            ExecutarComFreeze(() =>
            {
                try
                {
                    if (_grupo == null || _grupo.Linhas == null) return;

                    string filtro = "TODOS";
                    if (cboFiltro.Selected != null)
                        filtro = cboFiltro.Selected.Value;

                    List<LinhaImportacao> linhasFiltradas;
                    switch (filtro)
                    {
                        case "SUCESSO":
                            linhasFiltradas = _grupo.Linhas.Where(l => l.Status == StatusLinha.Sucesso).ToList();
                            break;
                        case "ERRO":
                            linhasFiltradas = _grupo.Linhas.Where(l => l.Status == StatusLinha.Erro).ToList();
                            break;
                        case "PENDENTE":
                            linhasFiltradas = _grupo.Linhas.Where(l => l.Status == StatusLinha.Pendente).ToList();
                            break;
                        default:
                            linhasFiltradas = _grupo.Linhas;
                            break;
                    }

                    var dt = GetDataTable();
                    if (dt == null) return;

                    dt.Rows.Clear();

                    foreach (var linha in OrdenarLinhas(linhasFiltradas))
                    {
                        AdicionarLinhaDataTable(dt, linha);
                    }

                    oGrid.Clear();
                    oGrid.LoadFromDataSource();
                    AplicarEstiloMatrixOtimizado();

                    AtualizarInterface();

                    MostrarStatus($"Filtro aplicado: {linhasFiltradas.Count} registro(s)");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erro em CboFiltro_ComboSelectAfter: {ex.Message}");
                }
            });
        }

        private void BtnProcessar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (_processamentoEmAndamento || _totalSelecionadas == 0)
            {
                if (_totalSelecionadas == 0)
                    MostrarMensagem("Não há documentos selecionados para processar.");
                return;
            }

            if (ConfirmarAcao($"Confirma o processamento de {_totalSelecionadas} documento(s)?\n\nValor total: {_valorTotalSelecionado:C}"))
            {
                IniciarProcessamento();
            }
        }

        private void BtnExportarErros_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            ExportarErros();
        }

        private void BtnVoltar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (!string.IsNullOrEmpty(_formOriginId))
                FormManager.TrazerParaFrente(_formOriginId);

            UIAPIRawForm.Close();
        }

        private void BtnFechar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            UIAPIRawForm.Close();
        }

        #endregion

        #region Form Events

        public override void OnInitializeFormEvents()
        {
            LoadAfter += Form_LoadAfter;
            CloseBefore += Form_CloseBefore;
            ResizeAfter += Form_ResizeAfter;
        }

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {
            UIAPIRawForm.Title = "Processamento de NFS-e em Lote - Resultado";
            FormManager.RegistrarFormulario(UIAPIRawForm.UniqueID, "Resultado Processamento", false, null, _grupo?.Code, 4);

            // Aguardar a Matrix carregar completamente antes de ajustar larguras
            Application.SBO_Application.Forms.ActiveForm.Freeze(true);
            try
            {
                if (oGrid != null)
                {
                    AjustarLarguraColunas();
                }
            }
            finally
            {
                Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }
        }

        private void Form_CloseBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Cancelar operações em andamento
                _cancellationTokenSource?.Cancel();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;

                //// Desconectar service layer
                //if (_serviceLayerClient != null)
                //{
                //    _serviceLayerClient.DisconnectAsync(CancellationToken.None).Wait(1000);
                //}

                // Limpar DataTable
                var dt = GetDataTable();
                dt?.Rows.Clear();

                // Limpar grid
                oGrid?.Clear();

                FormManager.RemoverFormulario(UIAPIRawForm.UniqueID);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro ao fechar: {ex.Message}");
            }

            // Limpar cache
            _ultimoStatusCache?.Clear();
            _zebraStripeAplicado = false;

            FormManager.RemoverFormulario(UIAPIRawForm.UniqueID);
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            AjustarLarguraColunas();
        }

        private void AjustarLarguraColunas()
        {
            try
            {
                if (oGrid == null || oGrid.RowCount == 0) return;

                var larguras = new Dictionary<string, int>
                {
                    ["#"] = 30,
                    ["Proc"] = 60,
                    ["Filial"] = 50,
                    ["CodCli"] = 80,
                    ["Cliente"] = 200,
                    ["CodItem"] = 80,
                    ["Descricao"] = 200,
                    ["Utiliz"] = 50,
                    ["DocEntry"] = 60,
                    ["DocNum"] = 60,
                    ["Status"] = 80,
                    ["CodImp"] = 70,
                    ["CodSeq"] = 60,
                    ["Condicao"] = 70,
                    ["ObsNF"] = 150,
                    ["TipoTrib"] = 60,
                    ["Valor"] = 80,
                    ["Mensagem"] = 300
                };

                foreach (var kvp in larguras)
                {
                    try
                    {
                        oGrid.Columns.Item(kvp.Key).Width = kvp.Value;
                    }
                    catch { /* Ignorar erro de coluna individual */ }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erro em AjustarLarguraColunas: {ex.Message}");
            }
        }

        #endregion

        #region Métodos Auxiliares

        private void ExportarErros()
        {
            if (_grupo == null || _grupo.Linhas == null)
            {
                MostrarMensagem("Não há dados para exportar.");
                return;
            }

            var erros = _grupo.Linhas.Where(l => l.Status == StatusLinha.Erro).ToList();

            if (erros.Count == 0)
            {
                MostrarMensagem("Não há erros para exportar.");
                return;
            }

            try
            {
                string caminhoArquivo = string.Empty;

                // Executar em thread STA para o diálogo funcionar corretamente
                Thread thread = new Thread(() =>
                {
                    // Criar dummy form para garantir que o diálogo apareça
                    using (System.Windows.Forms.Form dummyForm = new System.Windows.Forms.Form())
                    {
                        // Configurar dummy form
                        dummyForm.TopMost = true;
                        dummyForm.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                        dummyForm.ShowInTaskbar = false;
                        dummyForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                        dummyForm.Size = new System.Drawing.Size(1, 1);
                        dummyForm.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        dummyForm.Location = new System.Drawing.Point(-1000, -1000);
                        dummyForm.Opacity = 0;
                        
                        // Mostrar o form para criar handle válido
                        dummyForm.Show();
                        dummyForm.BringToFront();
                        
                        using (System.Windows.Forms.SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog())
                        {
                            // Configurar o diálogo
                            saveDialog.Filter = "Arquivo Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*";
                            saveDialog.Title = "Salvar relatório de erros NFS-e";
                            saveDialog.DefaultExt = "xlsx";
                            saveDialog.AddExtension = true;
                            saveDialog.OverwritePrompt = true;

                            // Sugerir nome do arquivo
                            string nomeArquivo = $"Erros_NFS_{_grupo.Nome}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                            saveDialog.FileName = nomeArquivo;

                            // Definir diretório inicial
                            saveDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                            // Usar dummy form como parent
                            if (saveDialog.ShowDialog(dummyForm) == System.Windows.Forms.DialogResult.OK)
                            {
                                caminhoArquivo = saveDialog.FileName;
                            }
                        }
                        
                        dummyForm.Close();
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                // Se o usuário cancelou, sair
                if (string.IsNullOrEmpty(caminhoArquivo))
                    return;

                // Exportar para o caminho escolhido
                var exportService = new ExcelExportService();
                string arquivo = exportService.ExportarErros(_grupo, erros, caminhoArquivo);

                MostrarMensagem($"Exportação concluída!\n\n📁 {arquivo}\n📊 {erros.Count} erro(s) exportado(s)");

                // Abrir o explorador de arquivos mostrando o arquivo
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{arquivo}\"");
            }
            catch (Exception ex)
            {
                MostrarErro($"Erro ao exportar: {ex.Message}");
            }
        }

        private List<LinhaImportacao> ObterLinhasSelecionadas()
        {
            var selecionadas = new List<LinhaImportacao>();

            if (_grupo == null || _grupo.Linhas == null) return selecionadas;

            var dt = GetDataTable();
            if (dt == null) return selecionadas;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.GetValue("Proc", i).ToString() == "Y")
                {
                    var code = dt.GetValue("Code", i).ToString();
                    var linha = _grupo.Linhas.FirstOrDefault(l => l.Code == code);

                    if (linha != null && linha.Status == StatusLinha.Pendente)
                        selecionadas.Add(linha);
                }
            }

            return selecionadas;
        }

        private DataTable GetDataTable()
        {
            try
            {
                if (UIAPIRawForm == null || UIAPIRawForm.DataSources == null || UIAPIRawForm.DataSources.DataTables == null)
                    return null;

                return UIAPIRawForm.DataSources.DataTables.Item("dtResult");
            }
            catch
            {
                return null;
            }
        }

        private void ExecutarComFreeze(Action acao)
        {
            if (UIAPIRawForm == null) return;

            try
            {
                // Mostrar mensagem antes de congelar
                Application.SBO_Application.StatusBar.SetText("Processando...", 
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                
                UIAPIRawForm.Freeze(true);
                acao();
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private string Truncar(string texto, int max)
        {
            if (string.IsNullOrEmpty(texto)) return "";
            return texto.Length <= max ? texto : texto.Substring(0, max - 3) + "...";
        }

        private void MostrarMensagem(string mensagem)
        {
            Application.SBO_Application.MessageBox(mensagem, 1, "Ok");
        }

        private void MostrarErro(string erro)
        {
            Application.SBO_Application.SetStatusBarMessage(erro, BoMessageTime.bmt_Short, true);
        }

        private void MostrarStatus(string status)
        {
            Application.SBO_Application.SetStatusBarMessage(status, BoMessageTime.bmt_Short, false);
        }

        private bool ConfirmarAcao(string mensagem)
        {
            return Application.SBO_Application.MessageBox(mensagem, 2, "Sim", "Não") == 1;
        }

        private void SelecionarTodasPendentes(bool selecionar)
        {
            ExecutarComFreeze(() =>
            {
                try
                {
                    if (_grupo == null || _grupo.Linhas == null) return;

                    var dt = GetDataTable();
                    if (dt == null) return;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var code = dt.GetValue("Code", i).ToString();
                        var linha = _grupo.Linhas.FirstOrDefault(l => l.Code == code);

                        if (linha != null && linha.Status == StatusLinha.Pendente)
                        {
                            dt.SetValue("Proc", i, selecionar ? "Y" : "N");
                        }
                    }

                    oGrid.LoadFromDataSource();
                    RecalcularTotais();
                    AtualizarStatus();
                    AtualizarBotoes();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erro em SelecionarTodasPendentes: {ex.Message}");
                }
            });
        }

        #endregion
    }
}