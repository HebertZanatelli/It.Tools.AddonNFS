using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using SAPbouiCOM.Framework;

namespace ItTech.Tool.AddonNFS.Utils
{
    /// <summary>
    /// Helper centralizado para gerenciar ProgressBar de forma thread-safe
    /// Pode ser usado em qualquer formulário do addon
    /// </summary>
    public class ProgressBarHelper : IDisposable
    {
        private static ProgressBarHelper _instance;
        private static readonly object _lockInstance = new object();

        private SAPbouiCOM.Application _sboApp;
        private SAPbouiCOM.ProgressBar _progressBar;
        private readonly object _lockProgressBar = new object();

        // Sistema de fila para thread-safe
        private System.Timers.Timer _timerUI;
        private Queue<Action> _filaAcoesUI = new Queue<Action>();
        private readonly object _lockFila = new object();

        // Estado
        private bool _isProgressBarActive = false;
        private int _currentValue = 0;
        private int _maxValue = 100;
        private string _lastError = "";

        #region Singleton Pattern

        public static ProgressBarHelper Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lockInstance)
                    {
                        if (_instance == null)
                        {
                            _instance = new ProgressBarHelper();
                        }
                    }
                }
                return _instance;
            }
        }

        private ProgressBarHelper()
        {
            _sboApp = SAPbouiCOM.Framework.Application.SBO_Application;
            IniciarTimerUI();
        }

        #endregion

        #region Métodos Públicos

        /// <summary>
        /// Cria uma nova progress bar (thread-safe)
        /// </summary>
        public void Criar(string texto, int maximo, bool podeSerCancelada = false)
        {
            EnfileirarAcaoUI(() =>
            {
                lock (_lockProgressBar)
                {
                    try
                    {
                        // Fechar anterior se existir
                        FecharInterno();

                        _progressBar = _sboApp.StatusBar.CreateProgressBar(texto, maximo, podeSerCancelada);
                        _isProgressBarActive = true;
                        _currentValue = 0;
                        _maxValue = maximo;
                    }
                    catch (Exception ex)
                    {
                        _lastError = ex.Message;
                        // Se falhar por "resource is occupied", tentar novamente após delay
                        if (ex.Message.Contains("resource is occupied"))
                        {
                            Thread.Sleep(100);
                            try
                            {
                                _progressBar = _sboApp.StatusBar.CreateProgressBar(texto, maximo, podeSerCancelada);
                                _isProgressBarActive = true;
                                _currentValue = 0;
                                _maxValue = maximo;
                            }
                            catch { }
                        }
                    }
                }
            });
        }

        /// <summary>
        /// Atualiza o valor e/ou texto da progress bar
        /// </summary>
        public void Atualizar(int valor, string texto = null)
        {
            EnfileirarAcaoUI(() =>
            {
                lock (_lockProgressBar)
                {
                    try
                    {
                        if (_progressBar != null && _isProgressBarActive)
                        {
                            // Garantir que não ultrapasse o máximo
                            if (valor > _maxValue)
                                valor = _maxValue;

                            _progressBar.Value = valor;
                            _currentValue = valor;

                            if (!string.IsNullOrEmpty(texto))
                                _progressBar.Text = texto;
                        }
                    }
                    catch
                    {
                        // Ignorar erros de atualização
                    }
                }
            });
        }

        /// <summary>
        /// Incrementa o valor da progress bar
        /// </summary>
        public void Incrementar(int incremento = 1, string texto = null)
        {
            int novoValor = _currentValue + incremento;
            Atualizar(novoValor, texto);
        }

        /// <summary>
        /// Atualiza com percentual
        /// </summary>
        public void AtualizarPercentual(double percentual, string texto = null)
        {
            int valor = (int)Math.Round((_maxValue * percentual) / 100.0);
            Atualizar(valor, texto);
        }

        /// <summary>
        /// Fecha a progress bar
        /// </summary>
        public void Fechar()
        {
            EnfileirarAcaoUI(() =>
            {
                lock (_lockProgressBar)
                {
                    FecharInterno();
                }
            });
        }

        /// <summary>
        /// Executa uma ação com progress bar
        /// </summary>
        public void ExecutarComProgress(string texto, int totalPassos, Action<Action<int, string>> acao)
        {
            try
            {
                Criar(texto, totalPassos);

                // Ação que atualiza o progresso
                Action<int, string> atualizador = (valor, msg) =>
                {
                    Atualizar(valor, msg);
                };

                // Executar a ação
                acao(atualizador);
            }
            finally
            {
                // Pequeno delay para mostrar 100%
                Thread.Sleep(300);
                Fechar();
            }
        }

        /// <summary>
        /// Executa uma ação assíncrona com progress bar
        /// </summary>
        public void ExecutarAsyncComProgress(string texto, int totalPassos, Action<Action<int, string>> acao)
        {
            Thread thread = new Thread(() =>
            {
                ExecutarComProgress(texto, totalPassos, acao);
            });
            thread.IsBackground = true;
            thread.Start();
        }

        #endregion

        #region Métodos Privados

        private void IniciarTimerUI()
        {
            if (_timerUI == null)
            {
                _timerUI = new System.Timers.Timer(100);
                _timerUI.Elapsed += ProcessarFilaUI;
                _timerUI.Start();
            }
        }

        private void EnfileirarAcaoUI(Action acao)
        {
            lock (_lockFila)
            {
                _filaAcoesUI.Enqueue(acao);
            }
        }

        private void ProcessarFilaUI(object sender, ElapsedEventArgs e)
        {
            List<Action> acoesParaProcessar = new List<Action>();

            lock (_lockFila)
            {
                // OTIMIZADO: Pegar até 3 ações por vez para processar em lote
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
                    acao.Invoke();
                }
                catch (Exception ex)
                {
                    _lastError = ex.Message;
                }
            }
        }

        private void FecharInterno()
        {
            try
            {
                if (_progressBar != null)
                {
                    _progressBar.Stop();
                    Marshal.ReleaseComObject(_progressBar);
                    _progressBar = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch { }
            finally
            {
                _isProgressBarActive = false;
                _currentValue = 0;
            }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            lock (_lockProgressBar)
            {
                FecharInterno();
            }

            if (_timerUI != null)
            {
                _timerUI.Stop();
                _timerUI.Dispose();
                _timerUI = null;
            }

            lock (_lockFila)
            {
                _filaAcoesUI.Clear();
            }
        }

        #endregion

        #region Propriedades

        public bool IsActive => _isProgressBarActive;
        public int CurrentValue => _currentValue;
        public int MaxValue => _maxValue;
        public string LastError => _lastError;

        #endregion
    }

    /// <summary>
    /// Extension methods para facilitar o uso
    /// </summary>
    public static class ProgressBarExtensions
    {
        /// <summary>
        /// Executa uma lista de tarefas com progress bar
        /// </summary>
        public static void ExecutarListaComProgress<T>(this IList<T> lista, string textoInicial,
            Action<T, int> processarItem)
        {
            ProgressBarHelper.Instance.ExecutarComProgress(textoInicial, lista.Count, (atualizador) =>
            {
                for (int i = 0; i < lista.Count; i++)
                {
                    processarItem(lista[i], i);
                    atualizador(i + 1, $"Processando item {i + 1} de {lista.Count}...");
                }
            });
        }

        /// <summary>
        /// Executa uma lista de tarefas assincronamente com progress bar
        /// </summary>
        public static void ExecutarListaAsyncComProgress<T>(this IList<T> lista, string textoInicial,
            Action<T, int> processarItem)
        {
            ProgressBarHelper.Instance.ExecutarAsyncComProgress(textoInicial, lista.Count, (atualizador) =>
            {
                for (int i = 0; i < lista.Count; i++)
                {
                    processarItem(lista[i], i);
                    atualizador(i + 1, $"Processando item {i + 1} de {lista.Count}...");
                }
            });
        }
    }
}