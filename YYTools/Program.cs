using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YYTools
{
    static class Program
    {
        private static AsyncTaskManager _taskManager = new AsyncTaskManager();
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Logger.LogInfo("开始启动程序");
                // 基本设置
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 显示启动信息
                // MessageBox.Show("正在启动YY工具...", "启动中", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("new MatchForm begin");
                // 先显示主窗体，保证快速可见
                var mainForm = new MatchForm();
                Logger.LogInfo("new MatchForm end");
                
                Logger.LogInfo("StartupProgressForm.ShowStartupProgress()");
                // 启动后台预热（解析已打开的Excel/WPS，填充缓存），并显示小型启动进度窗体
                // var progressForm = StartupProgressForm.ShowStartupProgress();
                // var cts = new System.Threading.CancellationTokenSource();
                // var progress = new Progress<YYTools.TaskProgress>(p => { try { progressForm?.UpdateProgress(p.Percentage, p.Message); } catch { } });
                //
                // _taskManager.StartBackgroundTask(
                //     taskName: "StartupWarmup",
                //     taskFactory: async (token, reporter) =>
                //     {
                //         // 桥接到统一的进度对象
                //         await AsyncStartupManager.WarmUpAsync(cts.Token, progress);
                //         try { progressForm?.CompleteStartup(true, ""); } catch { }
                //     },
                //     allowMultiple: false
                // );
                Logger.LogInfo("Application.Run(mainForm) begin");
                Application.Run(mainForm);
                Logger.LogInfo("Application.Run(mainForm) end");
            }
            catch (Exception ex)
            {
                // 显示详细错误信息
                string errorMessage = $"程序启动失败！\n\n" +
                                    $"错误类型: {ex.GetType().Name}\n" +
                                    $"错误信息: {ex.Message}\n\n" +
                                    $"堆栈跟踪:\n{ex.StackTrace}\n\n" +
                                    $"请将此信息发送给开发者。";
                
                MessageBox.Show(errorMessage, "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                // 尝试写入错误到文件
                try
                {
                    string errorLog = $"启动时间: {DateTime.Now}\n" +
                                     $"错误类型: {ex.GetType().Name}\n" +
                                     $"错误信息: {ex.Message}\n" +
                                     $"堆栈跟踪: {ex.StackTrace}\n" +
                                     $"操作系统: {Environment.OSVersion}\n" +
                                     $".NET版本: {Environment.Version}\n" +
                                     $"工作目录: {Environment.CurrentDirectory}\n";
                    
                    System.IO.File.WriteAllText("startup_error.log", errorLog);
                    MessageBox.Show("错误信息已保存到 startup_error.log 文件", "错误已保存", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    // 忽略保存错误日志失败
                }
                
                // 确保程序不会静默退出
                Application.Exit();
            }
        }
    }
}