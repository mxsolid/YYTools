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
                // 基本设置
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 显示启动信息
                MessageBox.Show("正在启动YY工具...", "启动中", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 快速启动主窗体
                var mainForm = new MatchForm();
                
                // 异步加载Excel文件，不阻塞启动
                Task.Run(() =>
                {
                    try
                    {
                        Thread.Sleep(100); // 短暂延迟，让主窗体先显示
                        mainForm.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                mainForm.InitializeExcelFilesAsync();
                            }
                            catch (Exception ex)
                            {
                                // 忽略Excel加载错误，不影响程序启动
                                System.Diagnostics.Debug.WriteLine($"Excel文件加载失败: {ex.Message}");
                            }
                        }));
                    }
                    catch
                    {
                        // 忽略异步加载错误
                    }
                });
                
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
                    System.IO.File.WriteAllText("startup_error.log", 
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 启动错误:\n{ex}\n\n堆栈跟踪:\n{ex.StackTrace}"); 
                } 
                catch { }
                
                // 确保程序不会静默退出
                Application.Exit();
            }
        }
    }
}