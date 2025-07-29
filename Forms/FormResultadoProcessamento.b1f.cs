using ItTech.Tool.AddonNFS.Controllers;
using ItTech.Tool.AddonNFS.Models;
using ItTech.Tool.AddonNFS.Services;
using ItTech.Tool.AddonNFS.Utils;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Application = SAPbouiCOM.Framework.Application;

namespace ItTech.Tool.AddonNFS.Forms
{
    /// <summary>
    /// Formulário de resultado e processamento de NFS-e em lote
    /// </summary>
    [FormAttribute("ITTECH_NFS_RESULT", "Forms/FormResultadoProcessamento.b1f")]
    class FormResultadoProcessamento : UserFormBase
    {
        #region Campos

        private GrupoLote _grupo;
        private ProcessamentoNFSController _processamentoController;
        private ServiceLayerInvoiceClient _serviceLayerClient;
        private bool _processamentoEmAndamento = false;
        private bool _atualizandoSelecao = false;
        private CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        private string _formOriginId;

        // Controles do formulário
        private StaticText LblTitulo;
        private StaticText LblStatus;
        private EditText TxtGrupo;
        private EditText TxtDtLan;
        private EditText TxtDtDoc;
        private ComboBox CmbFiltro;
        private Button BtnProcessar;
        private Button BtnExportarErros;
        private Button BtnVoltar;
        private Button BtnFechar;

        // Controle de seleção
        private Dictionary<string, bool> _linhasSelecionadas = new Dictionary<string, bool>();
        private int _totalSelecionadas = 0;
        private decimal _valorTotalSelecionado = 0;

        #endregion

        #region Construtores

        public FormResultadoProcessamento()
        {
            // Construtor padrão
        }

        public FormResultadoProcessamento(string formOriginId)
        {
            _formOriginId = formOriginId;
        }

        #endregion

        #region Inicialização

        public override void OnInitializeComponent()
        {
            this.LblTitulo = ((SAPbouiCOM.StaticText)(this.GetItem("lblTitulo").Specific));
            this.LblStatus = ((SAPbouiCOM.StaticText)(this.GetItem("lblStatus").Specific));
            this.TxtGrupo = ((SAPbouiCOM.EditText)(this.GetItem("txtGrp").Specific));
            this.TxtDtLan = ((SAPbouiCOM.EditText)(this.GetItem("txtDtLan").Specific));
            this.TxtDtDoc = ((SAPbouiCOM.EditText)(this.GetItem("txtDtDoc").Specific));
            this.CmbFiltro = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbFiltro").Specific));
            this.BtnProcessar = ((SAPbouiCOM.Button)(this.GetItem("btnProc").Specific));
            this.BtnProcessar.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.BtnProcessar_ClickBefore);
            this.BtnExportarErros = ((SAPbouiCOM.Button)(this.GetItem("btnExpErr").Specific));
            this.BtnVoltar = ((SAPbouiCOM.Button)(this.GetItem("btnVoltar").Specific));
            this.BtnFechar = ((SAPbouiCOM.Button)(this.GetItem("btnFechar").Specific));
            //    Eventos essenciais
            this.CmbFiltro.ComboSelectAfter += this.CmbFiltro_ComboSelectAfter;
            //   this.MatrixResultados.LinkPressedAfter += this.Matrix_LinkPressedAfter;
            //    this.MatrixResultados.ComboSelectAfter += this.Matrix_ComboSelectAfter;
            this.BtnProcessar.ClickBefore += this.BtnProcessar_ClickBefore;
            this.BtnExportarErros.ClickBefore += this.BtnExportarErros_ClickBefore;
            this.BtnVoltar.ClickBefore += this.BtnVoltar_ClickBefore;
            this.BtnFechar.ClickBefore += this.BtnFechar_ClickBefore;
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_1").Specific));
            
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            
            this.CloseBefore += Form_CloseBefore;
            this.ResizeAfter += Form_ResizeAfter;
        }

        private void OnCustomInitialize()
        {
            try
            {
                var company = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                _serviceLayerClient = new ServiceLayerInvoiceClient();
                _processamentoController = new ProcessamentoNFSController(company, _serviceLayerClient);
                System.Threading.Thread.Sleep(100); // Pequeno delay
                this.UIAPIRawForm.Height = 550; // Menor que 600 para forçar scroll
                this.UIAPIRawForm.Update();
                // ConfigurarDataTable();
                // ConfigurarMatrix();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao inicializar: {ex.Message}", 1, "Ok", "", "");
            }
        }

        #endregion

        #region SetGrupo - Entrada Principal

        public void SetGrupo(GrupoLote grupo)
        {
            if (grupo == null)
            {
                Application.SBO_Application.MessageBox("Grupo não pode ser nulo.", 1, "Ok", "", "");
                return;
            }

            _grupo = grupo;
            CarregarDados();
        }

        #endregion

        #region Configuração DataTable e Matrix

        private void ConfigurarDataTable()
        {
            var dt = UIAPIRawForm.DataSources.DataTables.Add("dtResult");

            // Adicionar todas as colunas conforme o XML SEM NumNF
            dt.Columns.Add("Linha", BoFieldsType.ft_Integer);
            dt.Columns.Add("Processar", BoFieldsType.ft_AlphaNumeric, 1);
            dt.Columns.Add("Filial", BoFieldsType.ft_Integer);
            dt.Columns.Add("CodCli", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("NomeCli", BoFieldsType.ft_AlphaNumeric, 100);
            dt.Columns.Add("CodItem", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("DescItem", BoFieldsType.ft_AlphaNumeric, 100);
            dt.Columns.Add("Utilizacao", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("CodImposto", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("SeqNF", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("CondPagto", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("Valor", BoFieldsType.ft_Float);
            dt.Columns.Add("DocEntry", BoFieldsType.ft_Integer);
            dt.Columns.Add("DocNum", BoFieldsType.ft_Integer);
            dt.Columns.Add("Status", BoFieldsType.ft_AlphaNumeric, 50);
            dt.Columns.Add("TipoTrib", BoFieldsType.ft_AlphaNumeric, 50);
            // NumNF REMOVIDA
            dt.Columns.Add("ObsNF", BoFieldsType.ft_AlphaNumeric, 254);
            dt.Columns.Add("Mensagem", BoFieldsType.ft_AlphaNumeric, 254);
            dt.Columns.Add("Code", BoFieldsType.ft_AlphaNumeric, 50); // Campo interno para referência
        }

        private void ConfigurarMatrix()
        {
            //// Bind de todas as colunas SEM NumNF
            //MatrixResultados.Columns.Item("Linha").DataBind.Bind("dtResult", "Linha");
            //MatrixResultados.Columns.Item("Processar").DataBind.Bind("dtResult", "Processar");
            //MatrixResultados.Columns.Item("Filial").DataBind.Bind("dtResult", "Filial");
            //MatrixResultados.Columns.Item("CodCli").DataBind.Bind("dtResult", "CodCli");
            //MatrixResultados.Columns.Item("NomeCli").DataBind.Bind("dtResult", "NomeCli");
            //MatrixResultados.Columns.Item("CodItem").DataBind.Bind("dtResult", "CodItem");
            //MatrixResultados.Columns.Item("DescItem").DataBind.Bind("dtResult", "DescItem");
            //MatrixResultados.Columns.Item("Utilizacao").DataBind.Bind("dtResult", "Utilizacao");
            //MatrixResultados.Columns.Item("CodImposto").DataBind.Bind("dtResult", "CodImposto");
            //MatrixResultados.Columns.Item("SeqNF").DataBind.Bind("dtResult", "SeqNF");
            //MatrixResultados.Columns.Item("CondPagto").DataBind.Bind("dtResult", "CondPagto");
            //MatrixResultados.Columns.Item("Valor").DataBind.Bind("dtResult", "Valor");
            //MatrixResultados.Columns.Item("DocEntry").DataBind.Bind("dtResult", "DocEntry");
            //MatrixResultados.Columns.Item("DocNum").DataBind.Bind("dtResult", "DocNum");
            //MatrixResultados.Columns.Item("Status").DataBind.Bind("dtResult", "Status");
            //MatrixResultados.Columns.Item("TipoTrib").DataBind.Bind("dtResult", "TipoTrib");
            //// NumNF REMOVIDA - NÃO FAZER BIND
            //MatrixResultados.Columns.Item("ObsNF").DataBind.Bind("dtResult", "ObsNF");
            //MatrixResultados.Columns.Item("Mensagem").DataBind.Bind("dtResult", "Mensagem");

            //// Configurar LinkedButtons
            //var lbCodCli = (LinkedButton)MatrixResultados.Columns.Item("CodCli").ExtendedObject;
            //lbCodCli.LinkedObject = BoLinkedObject.lf_BusinessPartner;

            //var lbCodItem = (LinkedButton)MatrixResultados.Columns.Item("CodItem").ExtendedObject;
            //lbCodItem.LinkedObject = BoLinkedObject.lf_Items;

            //var lbDocEntry = (LinkedButton)MatrixResultados.Columns.Item("DocEntry").ExtendedObject;
            //lbDocEntry.LinkedObject = BoLinkedObject.lf_Invoice;


        }

        #endregion

        #region Carregamento de Dados

        private void CarregarDados()
        {
            try
            {
                UIAPIRawForm.Freeze(true);

                // Preencher campos do cabeçalho
                TxtGrupo.Value = _grupo.Nome;
                TxtDtLan.Value = _grupo.DataLancamento.ToString("dd/MM/yyyy");
                TxtDtDoc.Value = _grupo.DataDocumento.ToString("dd/MM/yyyy");

                // Carregar linhas se necessário
                if (_grupo.Linhas == null || _grupo.Linhas.Count == 0)
                {
                    var grupoController = new GrupoLoteController(
                        (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany()
                    );
                    _grupo.Linhas = grupoController.ObterLinhasGrupo(_grupo.Code);
                }

                // Limpar seleções
                _linhasSelecionadas.Clear();
                _totalSelecionadas = 0;
                _valorTotalSelecionado = 0;

                //  CarregarMatrix();
                //  AtualizarInterface();
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void CarregarMatrix()
        {
            var dt = UIAPIRawForm.DataSources.DataTables.Item("dtResult");
            dt.Rows.Clear();

            foreach (var linha in _grupo.Linhas)
            {
                dt.Rows.Add();
                int index = dt.Rows.Count - 1;

                // Primeira coluna: número da linha
                dt.SetValue("Linha", index, linha.NumeroLinha);

                // Segunda coluna: Processar - lógica principal
                bool podeProcessar = linha.Status == StatusLinha.Pendente;
                string valorProcessar = podeProcessar ? "S" : "N";
                dt.SetValue("Processar", index, valorProcessar);

                // Inicializar seleção
                if (podeProcessar)
                {
                    _linhasSelecionadas[linha.Code] = true; // Por padrão, pendentes vêm marcadas
                    _totalSelecionadas++;
                    _valorTotalSelecionado += linha.Valor;
                }

                dt.SetValue("Filial", index, linha.Filial);
                dt.SetValue("CodCli", index, linha.CodigoCliente);
                dt.SetValue("NomeCli", index, TruncarTexto(linha.NomeCliente, 100));
                dt.SetValue("CodItem", index, linha.CodigoItem);
                dt.SetValue("DescItem", index, TruncarTexto(linha.DescricaoItem, 100));
                dt.SetValue("Utilizacao", index, linha.Utilizacao ?? "");
                dt.SetValue("CodImposto", index, linha.CodigoImposto ?? "");
                dt.SetValue("SeqNF", index, linha.CodSeq ?? "");
                dt.SetValue("CondPagto", index, linha.CondicaoPagamento ?? "");
                dt.SetValue("Valor", index, Convert.ToDouble(linha.Valor));
                dt.SetValue("DocEntry", index, linha.DocEntry ?? 0);
                dt.SetValue("DocNum", index, linha.DocNum ?? 0);

                // Status visual
                string statusText = linha.Status == StatusLinha.Sucesso ? "✓ Sucesso" :
                                   linha.Status == StatusLinha.Erro ? "✗ Erro" :
                                   "○ Pendente";
                dt.SetValue("Status", index, statusText);

                dt.SetValue("TipoTrib", index, linha.TipoTributacao ?? "");
                // NumNF REMOVIDA - NÃO SETAR VALOR
                dt.SetValue("ObsNF", index, TruncarTexto(linha.ObservacaoNF, 254));
                dt.SetValue("Mensagem", index, TruncarTexto(linha.MensagemErro, 254));
                dt.SetValue("Code", index, linha.Code);
            }

            //MatrixResultados.Clear();
            //MatrixResultados.LoadFromDataSource();

            // Desabilitar combo para linhas processadas
            // AtualizarEstadoCombos();

            // Usar ajuste manual ao invés de AutoResizeColumns
            AjustarLarguraColunas();
        }

        private void AjustarLarguraColunas()
        {
            try
            {
                UIAPIRawForm.Freeze(true);

                // Definir larguras fixas para cada coluna
                //MatrixResultados.Columns.Item("Linha").Width = 35;
                //MatrixResultados.Columns.Item("Processar").Width = 65;
                //MatrixResultados.Columns.Item("Filial").Width = 35;
                //MatrixResultados.Columns.Item("CodCli").Width = 80;
                //MatrixResultados.Columns.Item("NomeCli").Width = 180;
                //MatrixResultados.Columns.Item("CodItem").Width = 80;
                //MatrixResultados.Columns.Item("DescItem").Width = 150;
                //MatrixResultados.Columns.Item("Utilizacao").Width = 60;
                //MatrixResultados.Columns.Item("CodImposto").Width = 80;
                //MatrixResultados.Columns.Item("SeqNF").Width = 60;
                //MatrixResultados.Columns.Item("CondPagto").Width = 80;
                //MatrixResultados.Columns.Item("Valor").Width = 80;
                //MatrixResultados.Columns.Item("DocEntry").Width = 70;
                //MatrixResultados.Columns.Item("DocNum").Width = 70;
                //MatrixResultados.Columns.Item("Status").Width = 80;
                //MatrixResultados.Columns.Item("TipoTrib").Width = 70;
                //// NumNF REMOVIDA
                //MatrixResultados.Columns.Item("ObsNF").Width = 120;
                //MatrixResultados.Columns.Item("Mensagem").Width = 200;


            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        //private void AtualizarEstadoCombos()
        //{
        //    for (int i = 1; i <= MatrixResultados.RowCount; i++)
        //    {
        //        var statusCell = MatrixResultados.Columns.Item("Status").Cells.Item(i).Specific as EditText;
        //        var comboCell = MatrixResultados.Columns.Item("Processar").Cells.Item(i).Specific as ComboBox;

        //        if (statusCell != null && comboCell != null)
        //        {
        //            bool isPendente = statusCell.Value.Contains("Pendente");

        //            // Habilitar apenas se for pendente (permite edição)
        //            if (isPendente)
        //            {
        //                comboCell.Item.Enabled = true;
        //            }
        //            else
        //            {
        //                // Desabilitar se já foi processado (sucesso ou erro)
        //                comboCell.Item.Enabled = false;
        //                // Aplicar cor visual para desabilitados
        //                MatrixResultados.CommonSetting.SetCellBackColor(i, 2, 14737632); // Cinza claro na coluna Processar
        //            }
        //        }
        //    }
        //}

        #endregion

        #region Processamento

        private void IniciarProcessamento()
        {
            try
            {
                _processamentoEmAndamento = true;
                AtualizarBotoes();

                // Processar apenas linhas pendentes E marcadas como "Sim"
                var linhasParaProcessar = _grupo.Linhas
                    .Where(l => l.Status == StatusLinha.Pendente &&
                               _linhasSelecionadas.ContainsKey(l.Code) &&
                               _linhasSelecionadas[l.Code])
                    .ToList();

                if (linhasParaProcessar.Count == 0)
                {
                    Application.SBO_Application.MessageBox(
                        "Não há documentos selecionados para processar.\n\n" +
                        "Marque 'Sim' na coluna 'Processar' para os documentos desejados.",
                        1, "Ok", "", ""
                    );
                    _processamentoEmAndamento = false;
                    AtualizarInterface();
                    return;
                }

                // Usar ProgressBarHelper
                ProgressBarHelper.Instance.ExecutarComProgress(
                    $"Processando {linhasParaProcessar.Count} documentos...",
                    100,
                    (atualizador) =>
                    {
                        Task.Run(async () =>
                        {
                            await ProcessarLinhasAsync(linhasParaProcessar, atualizador);
                        }).Wait();
                    });
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro: {ex.Message}", 1, "Ok", "", "");
                _processamentoEmAndamento = false;
                AtualizarInterface();
            }
        }

        private async Task ProcessarLinhasAsync(List<LinhaImportacao> linhasParaProcessar, Action<int, string> atualizador)
        {
            var tempoInicio = DateTime.Now;
            int LOTE_SIZE = 25;

            using (var timeoutCancellation = new CancellationTokenSource(TimeSpan.FromMinutes(60)))
            using (var combinedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                _cancellationTokenSource.Token, timeoutCancellation.Token))
            {
                try
                {
                    // Validação
                    atualizador(5, "Validando dados SAP...");
                    var validacao = _processamentoController.ValidarDadosSAP(linhasParaProcessar);

                    if (!validacao.Valida && validacao.Erros.Count > 0)
                    {
                        string mErros = string.Join("\n", validacao.Erros.Take(5));
                        await ExecutarNaUIThread(() =>
                        {
                            Application.SBO_Application.MessageBox(
                                $"Erros encontrados:\n{mErros}\n\nCorrija os erros antes de continuar.",
                                1, "Ok", "", ""
                            );
                        });
                        return;
                    }

                    // Conexão Service Layer
                    atualizador(15, "Conectando ao Service Layer...");
                    bool conectado = await _serviceLayerClient.ConnectAsync(combinedCancellation.Token);

                    if (!conectado)
                    {
                        throw new Exception("Falha ao conectar com Service Layer");
                    }

                    atualizador(20, "Conectado! Iniciando processamento...");

                    // Processar em lotes
                    int totalLinhas = linhasParaProcessar.Count;
                    int processadas = 0;
                    int sucessoTotal = 0;
                    int erroTotal = 0;

                    for (int loteInicio = 0; loteInicio < totalLinhas; loteInicio += LOTE_SIZE)
                    {
                        combinedCancellation.Token.ThrowIfCancellationRequested();

                        int loteFim = Math.Min(loteInicio + LOTE_SIZE, totalLinhas);
                        var loteAtual = linhasParaProcessar.Skip(loteInicio).Take(loteFim - loteInicio).ToList();

                        int progresso = 20 + (int)((processadas / (double)totalLinhas) * 70);
                        atualizador(progresso, $"Processando {processadas + 1} a {loteFim} de {totalLinhas}...");

                        var resultadosLote = await Task.Run(() =>
                        {
                            return _processamentoController.ProcessarLinhas(
                                _grupo.Code,
                                loteAtual,
                                _grupo.DataLancamento,
                                _grupo.DataDocumento
                            );
                        }, combinedCancellation.Token);

                        // Atualizar UI parcialmente
                        await ExecutarNaUIThread(() =>
                        {
                            AtualizarInterfaceParcial(resultadosLote);
                        });

                        sucessoTotal += resultadosLote.Count(r => r.Sucesso);
                        erroTotal += resultadosLote.Count(r => !r.Sucesso);
                        processadas = loteFim;

                        if (loteInicio + LOTE_SIZE < totalLinhas)
                        {
                            await Task.Delay(100, combinedCancellation.Token);
                        }
                    }

                    atualizador(95, "Finalizando processamento...");

                    var tempoTotal = DateTime.Now - tempoInicio;

                    await ExecutarNaUIThread(() =>
                    {
                        Application.SBO_Application.SetStatusBarMessage(
                            $"✅ Processamento concluído: {sucessoTotal} sucessos, {erroTotal} erros em {tempoTotal.TotalSeconds:F1}s",
                            BoMessageTime.bmt_Long,
                            false
                        );

                        Application.SBO_Application.MessageBox(
                            $"Processamento concluído!\n\n" +
                            $"✅ Sucessos: {sucessoTotal}\n" +
                            $"❌ Erros: {erroTotal}\n" +
                            $"⏱️ Tempo: {tempoTotal.TotalSeconds:F1}s\n" +
                            $"📊 Total processado: {processadas} documentos",
                            1, "Ok", "", ""
                        );
                    });

                    atualizador(100, "Concluído!");
                }
                catch (OperationCanceledException)
                {
                    await ExecutarNaUIThread(() =>
                    {
                        Application.SBO_Application.MessageBox(
                            "⚠️ Processamento interrompido!",
                            1, "Ok", "", ""
                        );
                    });
                }
                catch (Exception ex)
                {
                    await ExecutarNaUIThread(() =>
                    {
                        Application.SBO_Application.MessageBox(
                            $"❌ Erro no processamento: {ex.Message}",
                            1, "Ok", "", ""
                        );
                    });
                }
                finally
                {
                    _processamentoEmAndamento = false;
                    await ExecutarNaUIThread(() =>
                    {
                        CarregarDados();
                    });
                }
            }
        }

        private void AtualizarInterfaceParcial(List<ResultadoProcessamento> resultadosLote)
        {
            try
            {
                if (UIAPIRawForm == null || resultadosLote == null) return;

                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtResult");

                foreach (var resultado in resultadosLote)
                {
                    // Encontrar linha pelo Code
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.GetValue("Code", i).ToString() == resultado.CodigoLinha)
                        {
                            // Atualizar campos
                            string statusText = resultado.Sucesso ? "✓ Sucesso" : "✗ Erro";
                            dt.SetValue("Status", i, statusText);
                            dt.SetValue("Mensagem", i, TruncarTexto(resultado.Mensagem, 254));
                            dt.SetValue("Processar", i, "N"); // Não pode mais processar

                            if (resultado.DocEntry.HasValue)
                            {
                                dt.SetValue("DocEntry", i, resultado.DocEntry.Value);
                                dt.SetValue("DocNum", i, resultado.DocNum ?? resultado.DocEntry.Value);
                            }

                            // Atualizar objeto
                            var linha = _grupo.Linhas.FirstOrDefault(l => l.Code == resultado.CodigoLinha);
                            if (linha != null)
                            {
                                linha.Status = resultado.Sucesso ? StatusLinha.Sucesso : StatusLinha.Erro;
                                linha.MensagemErro = resultado.Mensagem;
                                linha.DocEntry = resultado.DocEntry;
                                linha.DocNum = resultado.DocNum;

                                // Remover da seleção
                                _linhasSelecionadas.Remove(linha.Code);
                            }

                            break;
                        }
                    }
                }

                // Atualizar Matrix
                //MatrixResultados.LoadFromDataSource();
                //AtualizarEstadoCombos();
                AjustarLarguraColunas();

                // Atualizar status
                RecalcularTotais();
                AtualizarStatus();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao atualizar: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Interface e Status

        private void AtualizarInterface()
        {
            try
            {
                AtualizarStatus();
                AtualizarBotoes();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage($"Erro ao atualizar interface: {ex.Message}",
                    BoMessageTime.bmt_Short, true);
            }
        }

        private void AtualizarStatus()
        {
            var stats = CalcularEstatisticas();

            string textoStatus = $"Status: {stats.Total} processados | ✓ {stats.Sucessos} sucessos | ✕ {stats.Erros} erros | ⚠ {stats.Pendentes} pendentes";

            // Adicionar informação de seleção
            if (_totalSelecionadas > 0)
            {
                textoStatus += $" | 📌 {_totalSelecionadas} selecionadas ({_valorTotalSelecionado:C})";
            }

            if (LblStatus.Caption != textoStatus)
            {
                LblStatus.Caption = textoStatus;
            }
        }

        private void AtualizarBotoes()
        {
            var stats = CalcularEstatisticas();

            bool temSelecionadas = _totalSelecionadas > 0;
            bool temErros = stats.Erros > 0;

            BtnProcessar.Item.Enabled = temSelecionadas && !_processamentoEmAndamento;
            BtnExportarErros.Item.Enabled = temErros;
        }

        private (int Total, int Sucessos, int Erros, int Pendentes) CalcularEstatisticas()
        {
            if (_grupo?.Linhas == null) return (0, 0, 0, 0);

            int total = _grupo.Linhas.Count;
            int sucessos = _grupo.Linhas.Count(l => l.Status == StatusLinha.Sucesso);
            int erros = _grupo.Linhas.Count(l => l.Status == StatusLinha.Erro);
            int pendentes = _grupo.Linhas.Count(l => l.Status == StatusLinha.Pendente);

            return (total, sucessos, erros, pendentes);
        }

        private void RecalcularTotais()
        {
            _totalSelecionadas = 0;
            _valorTotalSelecionado = 0;

            foreach (var linha in _grupo.Linhas)
            {
                if (linha.Status == StatusLinha.Pendente &&
                    _linhasSelecionadas.ContainsKey(linha.Code) &&
                    _linhasSelecionadas[linha.Code])
                {
                    _totalSelecionadas++;
                    _valorTotalSelecionado += linha.Valor;
                }
            }
        }

        #endregion

        #region Eventos

        //private void Matrix_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        //{
        //    if (_atualizandoSelecao || pVal.ColUID != "Processar") return;

        //    try
        //    {
        //        _atualizandoSelecao = true;

        //        // Obter o código da linha
        //        var codeCell = MatrixResultados.Columns.Item("Code").Cells.Item(pVal.Row).Specific as EditText;
        //        var comboCell = MatrixResultados.Columns.Item("Processar").Cells.Item(pVal.Row).Specific as ComboBox;

        //        if (codeCell != null && comboCell != null)
        //        {
        //            string code = codeCell.Value;
        //            bool selecionado = comboCell.Selected?.Value == "S";

        //            // Atualizar dicionário
        //            _linhasSelecionadas[code] = selecionado;

        //            // Recalcular totais
        //            RecalcularTotais();

        //            // Atualizar interface
        //            AtualizarStatus();
        //            AtualizarBotoes();
        //        }
        //    }
        //    finally
        //    {
        //        _atualizandoSelecao = false;
        //    }
        //}

        private void CmbFiltro_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (_grupo == null || _processamentoEmAndamento) return;

            UIAPIRawForm.Freeze(true);
            try
            {
                int filtro = Convert.ToInt32(CmbFiltro.Selected.Value);
                List<LinhaImportacao> linhasFiltradas;

                switch (filtro)
                {
                    case 1: // Apenas Sucessos
                        linhasFiltradas = _grupo.Linhas.Where(l => l.Status == StatusLinha.Sucesso).ToList();
                        break;
                    case 2: // Apenas Erros
                        linhasFiltradas = _grupo.Linhas.Where(l => l.Status == StatusLinha.Erro).ToList();
                        break;
                    case 3: // Apenas Pendentes
                        linhasFiltradas = _grupo.Linhas.Where(l => l.Status == StatusLinha.Pendente).ToList();
                        break;
                    default: // Todos
                        linhasFiltradas = _grupo.Linhas;
                        break;
                }

                // Recarregar matrix com filtro
                var dt = UIAPIRawForm.DataSources.DataTables.Item("dtResult");
                dt.Rows.Clear();

                foreach (var linha in linhasFiltradas)
                {
                    dt.Rows.Add();
                    int index = dt.Rows.Count - 1;

                    // Primeira coluna: número da linha
                    dt.SetValue("Linha", index, linha.NumeroLinha);

                    // Segunda coluna: Processar
                    bool podeProcessar = linha.Status == StatusLinha.Pendente;
                    string valorProcessar = (podeProcessar && _linhasSelecionadas.ContainsKey(linha.Code) && _linhasSelecionadas[linha.Code]) ? "S" : "N";
                    dt.SetValue("Processar", index, valorProcessar);

                    dt.SetValue("Filial", index, linha.Filial);
                    dt.SetValue("CodCli", index, linha.CodigoCliente);
                    dt.SetValue("NomeCli", index, TruncarTexto(linha.NomeCliente, 100));
                    dt.SetValue("CodItem", index, linha.CodigoItem);
                    dt.SetValue("DescItem", index, TruncarTexto(linha.DescricaoItem, 100));
                    dt.SetValue("Utilizacao", index, linha.Utilizacao ?? "");
                    dt.SetValue("CodImposto", index, linha.CodigoImposto ?? "");
                    dt.SetValue("SeqNF", index, linha.CodSeq ?? "");
                    dt.SetValue("CondPagto", index, linha.CondicaoPagamento ?? "");
                    dt.SetValue("Valor", index, Convert.ToDouble(linha.Valor));
                    dt.SetValue("DocEntry", index, linha.DocEntry ?? 0);
                    dt.SetValue("DocNum", index, linha.DocNum ?? 0);

                    string statusText = linha.Status == StatusLinha.Sucesso ? "✓ Sucesso" :
                                       linha.Status == StatusLinha.Erro ? "✗ Erro" : "○ Pendente";
                    dt.SetValue("Status", index, statusText);

                    dt.SetValue("TipoTrib", index, linha.TipoTributacao ?? "");
                    dt.SetValue("ObsNF", index, TruncarTexto(linha.ObservacaoNF, 254));
                    dt.SetValue("Mensagem", index, TruncarTexto(linha.MensagemErro, 254));
                    dt.SetValue("Code", index, linha.Code);
                }

                //MatrixResultados.Clear();
                //MatrixResultados.LoadFromDataSource();
                //AtualizarEstadoCombos();
                AjustarLarguraColunas();
            }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void Matrix_LinkPressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            // Evento para quando o usuário clica em LinkedButton
            // O SAP B1 já abre automaticamente o documento vinculado
        }

        private void BtnProcessar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (_processamentoEmAndamento) return;

            if (_totalSelecionadas == 0)
            {
                Application.SBO_Application.MessageBox(
                    "Não há documentos selecionados para processar.\n\n" +
                    "Marque 'Sim' na coluna 'Processar' para os documentos desejados.",
                    1, "Ok", "", "");
                return;
            }

            string mensagem = $"Confirma o processamento de {_totalSelecionadas} documento(s) selecionado(s)?\n\n" +
                             $"Valor total: {_valorTotalSelecionado:C}";

            int resposta = Application.SBO_Application.MessageBox(mensagem, 2, "Sim", "Não", "");
            if (resposta == 1)
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

            // Voltar para formulário anterior
            if (!string.IsNullOrEmpty(_formOriginId))
            {
                FormManager.TrazerParaFrente(_formOriginId);
            }

            UIAPIRawForm.Close();
        }

        private void BtnFechar_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            UIAPIRawForm.Close();
        }

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {
            UIAPIRawForm.Title = "Processamento de NFS-e em Lote - Resultado";
            FormManager.RegistrarFormulario(UIAPIRawForm.UniqueID, "Resultado Processamento", false, null, _grupo?.Code, 4);
            this.UIAPIRawForm.Height = 600; // Menor que 699 original
            this.UIAPIRawForm.Update();
            // Aplicar ajuste de colunas após carregamento
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 500; // 500ms
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();

                //if (MatrixResultados.RowCount > 0)
                //{
                //    AjustarLarguraColunas();
                //}

                // Ajustar componentes após carregamento
                AjustarComponentesParaTamanhoJanela();
            };
            timer.Start();
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            AjustarComponentesParaTamanhoJanela();
        }

        private void AjustarComponentesParaTamanhoJanela()
        {
            try
            {
                UIAPIRawForm.Freeze(true);

                int formWidth = UIAPIRawForm.Width;
                int formHeight = UIAPIRawForm.Height;

                // Ajusta Matrix
                var matrix = this.GetItem("Item_1");
                if (matrix != null)
                {
                    matrix.Width = formWidth - matrix.Left - 21;
                    matrix.Height = formHeight - matrix.Top - 160;
                }

                // Ajusta label de progresso
                var lblProg = this.GetItem("lblProg");
                if (lblProg != null && matrix != null)
                {
                    lblProg.Top = matrix.Top + matrix.Height + 10;
                    lblProg.Width = matrix.Width;
                }

                // Ajusta posições dos botões
                int posYBotoes1 = lblProg != null ? lblProg.Top + 30 : formHeight - 110;
                int posYBotoes2 = posYBotoes1 + 28;

                // Primeira linha de botões
                var btnProc = this.GetItem("btnProc");
                if (btnProc != null)
                {
                    btnProc.Top = posYBotoes1;
                }

                var btnExpErr = this.GetItem("btnExpErr");
                if (btnExpErr != null)
                {
                    btnExpErr.Top = posYBotoes1;
                }

                // Segunda linha de botões
                var btnVoltar = this.GetItem("btnVoltar");
                if (btnVoltar != null)
                {
                    btnVoltar.Top = posYBotoes2;
                }

                // Botão Fechar - sempre no canto direito
                var btnFechar = this.GetItem("btnFechar");
                if (btnFechar != null)
                {
                    btnFechar.Top = posYBotoes2;
                    btnFechar.Left = formWidth - btnFechar.Width - 21;
                }

                // Ajustar combo de filtro se a janela for grande
                if (formWidth > 1200)
                {
                    var cmbFiltro = this.GetItem("cmbFiltro");
                    if (cmbFiltro != null)
                    {
                        cmbFiltro.Width = 200;
                        cmbFiltro.Left = formWidth - cmbFiltro.Width - 21;
                    }

                    var lblFiltro = this.GetItem("lblFiltro");
                    if (lblFiltro != null && cmbFiltro != null)
                    {
                        lblFiltro.Left = cmbFiltro.Left - 55;
                    }
                }
            }
            catch { }
            finally
            {
                UIAPIRawForm.Freeze(false);
            }
        }

        private void Form_CloseBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                _cancellationTokenSource?.Cancel();

                Task.Run(async () =>
                {
                    try
                    {
                        await _serviceLayerClient?.DisconnectAsync(CancellationToken.None);
                    }
                    catch { }
                });

                FormManager.RemoverFormulario(UIAPIRawForm.UniqueID);
            }
            catch { }
            finally
            {
                _cancellationTokenSource?.Dispose();
            }
        }

        #endregion

        #region Exportação de Erros

        private void ExportarErros()
        {
            try
            {
                if (_grupo == null || _grupo.Linhas == null)
                {
                    Application.SBO_Application.MessageBox("Não há dados para exportar.");
                    return;
                }

                var linhasComErro = _grupo.Linhas.Where(l => l.Status == StatusLinha.Erro).ToList();

                if (linhasComErro.Count == 0)
                {
                    Application.SBO_Application.MessageBox("Não há erros para exportar.", 1, "Ok", "", "");
                    return;
                }
                string caminhoArquivo = string.Empty;
                // Usar ExcelExportService para gerar planilha
                var exportService = new ExcelExportService();

                Thread thread = new Thread(() =>
                {
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

                        if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            caminhoArquivo = saveDialog.FileName;
                        }
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                // Se o usuário cancelou, sair
                if (string.IsNullOrEmpty(caminhoArquivo))
                    return;


                caminhoArquivo = exportService.ExportarErros(_grupo, linhasComErro, caminhoArquivo);

                Application.SBO_Application.MessageBox(
                    $"Exportação concluída!\n\n" +
                    $"📁 Arquivo salvo em:\n{caminhoArquivo}\n\n" +
                    $"📊 {linhasComErro.Count} erro(s) exportado(s)\n" +
                    $"✏️ Corrija os dados e importe novamente",
                    1, "Ok", "", ""
                );

                // Abrir pasta do arquivo
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{caminhoArquivo}\"");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox($"Erro ao exportar: {ex.Message}", 1, "Ok", "", "");
            }
        }

        #endregion

        #region Métodos Auxiliares

        /// <summary>
        /// Trunca string para tamanho máximo permitido
        /// </summary>
        private string TruncarTexto(string texto, int tamanhoMaximo = 254)
        {
            if (string.IsNullOrEmpty(texto))
                return "";

            if (texto.Length <= tamanhoMaximo)
                return texto;

            // Truncar e adicionar "..." no final
            return texto.Substring(0, tamanhoMaximo - 3) + "...";
        }

        private async Task ExecutarNaUIThread(Action acao)
        {
            await Task.Run(() =>
            {
                try
                {
                    acao?.Invoke();
                }
                catch { }
            });
        }

        #endregion

        private Matrix Matrix0;
    }
}