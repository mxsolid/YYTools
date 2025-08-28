using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading; // Added for Thread.Sleep

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

                // 快速启动主窗体，跳过所有复杂初始化
                var mainForm = new MatchForm();
                
                // 显示主窗体
                Application.Run(mainForm);
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
    }
}