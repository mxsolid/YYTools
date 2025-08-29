using System;
using System.Runtime.InteropServices;
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
                // TrySetPerMonitorV2DpiAwareness();

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);


                // 显示启动信息
                // MessageBox.Show("正在启动YY工具...", "启动中", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("new MatchForm begin");
                // 先显示主窗体，保证快速可见
                var mainForm = new MatchForm();
                Logger.LogInfo("new MatchForm end");

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
        
        
        static void TrySetPerMonitorV2DpiAwareness()
        {
            // 先尝试 SetProcessDpiAwarenessContext (Per-Monitor V2)
            if (SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
            {
                Log("SetProcessDpiAwarenessContext -> PER_MONITOR_AWARE_V2");
                return;
            }

            // 回退：SetProcessDpiAwareness (shcore.dll) -> PROCESS_PER_MONITOR_DPI_AWARE
            try
            {
                var result = SetProcessDpiAwareness(PROCESS_DPI_AWARENESS.PROCESS_PER_MONITOR_DPI_AWARE);
                if (result == 0) // S_OK
                {
                    Log("SetProcessDpiAwareness -> PROCESS_PER_MONITOR_DPI_AWARE");
                    return;
                }
            }
            catch { /* shcore.dll 可能不存在 */ }

            // 最后回退：SetProcessDPIAware (legacy)
            if (SetProcessDPIAware())
            {
                Log("SetProcessDPIAware -> success");
                return;
            }

            Log("Failed to set any DPI awareness API (falling back to default).");
        }

        static void Log(string msg)
        {
            try { System.Diagnostics.Debug.WriteLine("[DPI] " + msg); } catch { }
        }

        // P/Invoke declarations
        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

        // shcore.dll
        private enum PROCESS_DPI_AWARENESS
        {
            PROCESS_DPI_UNAWARE = 0,
            PROCESS_SYSTEM_DPI_AWARE = 1,
            PROCESS_PER_MONITOR_DPI_AWARE = 2
        }

        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(PROCESS_DPI_AWARENESS value);

        // legacy
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
    }
}