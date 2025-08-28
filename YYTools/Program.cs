using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YYTools
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 显示启动进度窗体
                var startupForm = StartupProgressForm.ShowStartupProgress();
                if (startupForm == null)
                {
                    // 如果启动进度窗体创建失败，直接启动主窗体
                    Application.Run(new MatchForm());
                    return;
                }

                // 创建异步启动管理器
                var startupManager = new AsyncStartupManager();
                
                // 绑定启动进度事件
                startupManager.ProgressReported += (s, e) =>
                {
                    if (startupForm != null && !startupForm.IsDisposed)
                    {
                        startupForm.UpdateProgress(e.Percentage, e.Message);
                    }
                };

                // 绑定启动完成事件
                startupManager.StartupCompleted += (s, e) =>
                {
                    if (startupForm != null && !startupForm.IsDisposed)
                    {
                        startupForm.CompleteStartup(e.Success, e.Message);
                    }

                    if (e.Success)
                    {
                        // 启动成功，显示主窗体
                        Application.Run(new MatchForm());
                    }
                    else
                    {
                        // 启动失败，退出应用程序
                        Application.Exit();
                    }
                };

                // 开始异步启动
                _ = startupManager.StartAsync();
            }
            catch (Exception ex)
            {
                // 记录错误并显示错误信息
                try
                {
                    Logger.LogError("应用程序启动失败", ex);
                }
                catch
                {
                    // 如果日志系统也失败，直接显示错误
                }

                MessageBox.Show($"应用程序启动失败：{ex.Message}\n\n请检查系统环境并重试。", 
                    "启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}