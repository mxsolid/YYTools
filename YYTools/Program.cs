using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading; // Added for Thread.Sleep

namespace YYTools
{
    static class Program
    {
        /// <summary>
        /// 自定义应用程序上下文：先显示启动进度窗体，启动完成后切换到主窗体。
        /// </summary>
        private class StartupApplicationContext : ApplicationContext
        {
            private StartupProgressForm _progressForm;
            private AsyncStartupManager _startupManager;
            private MatchForm _mainForm;

            public StartupApplicationContext()
            {
                try
                {
                    // 显示启动进度窗体
                    _progressForm = StartupProgressForm.ShowStartupProgress();

                    // 订阅进度事件
                    _startupManager = new AsyncStartupManager();
                    _startupManager.ProgressReported += (s, e) =>
                    {
                        try { _progressForm?.UpdateProgress(e.Percentage, e.Message); } catch { }
                    };
                    _startupManager.StartupCompleted += (s, e) =>
                    {
                        try { _progressForm?.CompleteStartup(e.Success, e.Message); } catch { }
                        OnStartupCompleted(e.Success);
                    };

                    // 异步启动
                    Task.Run(async () => { await _startupManager.StartAsync(); }).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"应用程序启动失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ExitThread();
                }
            }

            private void OnStartupCompleted(bool success)
            {
                if (!success)
                {
                    // 启动失败则退出
                    ExitThreadSafe();
                    return;
                }

                // 在UI线程切换到主窗体
                if (_progressForm != null)
                {
                    _progressForm.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            _mainForm = new MatchForm();
                            _mainForm.FormClosed += (s, e) => ExitThread();
                            _mainForm.Show();

                            // 关闭进度窗体
                            try { _progressForm.Close(); } catch { }
                            _progressForm = null;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"创建主窗体失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ExitThread();
                        }
                    }));
                }
                else
                {
                    // 回退：直接创建主窗体
                    _mainForm = new MatchForm();
                    _mainForm.FormClosed += (s, e) => ExitThread();
                    _mainForm.Show();
                }
            }

            private void ExitThreadSafe()
            {
                if (_progressForm != null)
                {
                    try { _progressForm.Close(); } catch { }
                    _progressForm = null;
                }
                ExitThread();
            }
        }
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                WriteLog("开始启动程序", LogLevel.Info);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 使用自定义应用程序上下文，确保单一消息循环并安全切换窗体
                var appContext = new StartupApplicationContext();
                Application.Run(appContext);
            }
            catch (Exception ex)
            {
                string errorMessage = $"程序启动失败！\n\n错误类型: {ex.GetType().Name}\n错误信息: {ex.Message}\n\n堆栈跟踪:\n{ex.StackTrace}\n\n请将此信息发送给开发者。";
                MessageBox.Show(errorMessage, "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                try 
                { 
                    System.IO.File.WriteAllText("startup_error.log", 
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 启动错误:\n{ex}\n\n堆栈跟踪:\n{ex.StackTrace}"); 
                } 
                catch { }
                Application.Exit();
            }
        }
        
        static void WriteLog(string message, LogLevel level) => MatchService.WriteLog($"[Program] {message}", level);

    }
}