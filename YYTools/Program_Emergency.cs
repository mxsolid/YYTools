using System;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 紧急修复版本 - 完全跳过Logger和复杂功能
    /// </summary>
    static class Program_Emergency
    {
        /// <summary>
        /// 应用程序的主入口点（紧急修复版本）
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                // 基本设置
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 显示启动信息
                MessageBox.Show("正在启动YY工具...", "启动中", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 直接创建并显示主窗体
                var mainForm = new MatchForm();
                
                // 显示启动成功信息
                MessageBox.Show("主窗体创建成功，正在显示...", "启动成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // 运行主窗体
                Application.Run(mainForm);
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