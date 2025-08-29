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
                
                // 启用DPI感知 - 必须在创建任何窗口之前调用
                bool dpiAwarenessSet = TrySetPerMonitorV2DpiAwareness();
                Logger.LogInfo($"DPI感知设置结果: {dpiAwarenessSet}");

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 显示启动信息
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
        
        /// <summary>
        /// 尝试设置Per-Monitor V2 DPI感知
        /// 返回是否成功设置
        /// </summary>
        static bool TrySetPerMonitorV2DpiAwareness()
        {
            try
            {
                // 方法1: 尝试 SetProcessDpiAwarenessContext (Per-Monitor V2) - Windows 10 1703+
                if (Environment.OSVersion.Version.Major >= 10 && Environment.OSVersion.Version.Build >= 15063)
                {
                    if (SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
                    {
                        Log("SetProcessDpiAwarenessContext -> PER_MONITOR_AWARE_V2 (成功)");
                        return true;
                    }
                    else
                    {
                        Log("SetProcessDpiAwarenessContext -> 失败");
                    }
                }

                // 方法2: 尝试 SetProcessDpiAwareness (shcore.dll) -> PROCESS_PER_MONITOR_DPI_AWARE
                try
                {
                    var result = SetProcessDpiAwareness(PROCESS_DPI_AWARENESS.PROCESS_PER_MONITOR_DPI_AWARE);
                    if (result == 0) // S_OK
                    {
                        Log("SetProcessDpiAwareness -> PROCESS_PER_MONITOR_DPI_AWARE (成功)");
                        return true;
                    }
                    else
                    {
                        Log($"SetProcessDpiAwareness -> 失败，错误代码: {result}");
                    }
                }
                catch (DllNotFoundException)
                {
                    Log("shcore.dll 不存在，跳过 SetProcessDpiAwareness");
                }
                catch (Exception ex)
                {
                    Log($"SetProcessDpiAwareness 异常: {ex.Message}");
                }

                // 方法3: 回退到 SetProcessDPIAware (legacy)
                if (SetProcessDPIAware())
                {
                    Log("SetProcessDPIAware -> 成功 (legacy)");
                    return true;
                }
                else
                {
                    Log("SetProcessDPIAware -> 失败");
                }

                Log("所有DPI感知API都失败了，使用系统默认设置");
                return false;
            }
            catch (Exception ex)
            {
                Log($"设置DPI感知时发生异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 记录DPI设置日志
        /// </summary>
        static void Log(string msg)
        {
            try 
            { 
                string logMessage = $"[DPI] {msg}";
                System.Diagnostics.Debug.WriteLine(logMessage);
                Logger.LogInfo(logMessage);
            } 
            catch 
            {
                // 如果Logger还没初始化，只输出到Debug
                System.Diagnostics.Debug.WriteLine($"[DPI] {msg}");
            } 
        }

        #region P/Invoke 声明

        // Windows 10 1703+ (Build 15063+)
        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

        // Windows 8.1+ (shcore.dll)
        private enum PROCESS_DPI_AWARENESS
        {
            PROCESS_DPI_UNAWARE = 0,           // 系统DPI感知
            PROCESS_SYSTEM_DPI_AWARE = 1,      // 系统DPI感知
            PROCESS_PER_MONITOR_DPI_AWARE = 2  // 每显示器DPI感知
        }

        [DllImport("shcore.dll", SetLastError = true)]
        private static extern int SetProcessDpiAwareness(PROCESS_DPI_AWARENESS value);

        // Windows Vista+ (legacy)
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDPIAware();

        #endregion
    }
}