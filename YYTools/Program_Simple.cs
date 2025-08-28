using System;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 简化的程序启动类（用于测试）
    /// </summary>
    static class Program_Simple
    {
        /// <summary>
        /// 应用程序的主入口点（简化版本）
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 直接启动主窗体，跳过异步启动
                var mainForm = new MatchForm();
                
                // 记录启动信息
                try
                {
                    Logger.LogInfo("应用程序启动（简化模式）");
                }
                catch
                {
                    // 忽略日志记录失败
                }
                
                // 运行主窗体
                Application.Run(mainForm);
            }
            catch (Exception ex)
            {
                // 显示错误信息
                MessageBox.Show($"应用程序启动失败：{ex.Message}\n\n请检查系统环境并重试。", 
                    "启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                // 确保程序不会静默退出
                Application.Exit();
            }
        }
    }
}