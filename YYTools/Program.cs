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

                // 先尝试直接启动主窗体，确保基本功能正常
                var mainForm = new MatchForm();
                
                // 尝试显示启动进度窗体
                StartupProgressForm startupForm = null;
                try
                {
                    startupForm = StartupProgressForm.ShowStartupProgress();
                }
                catch (Exception ex)
                {
                    // 如果启动进度窗体创建失败，记录错误但继续运行
                    try
                    {
                        Logger.LogWarning($"启动进度窗体创建失败: {ex.Message}");
                    }
                    catch
                    {
                        // 忽略日志记录失败
                    }
                }

                if (startupForm != null && !startupForm.IsDisposed)
                {
                    // 创建异步启动管理器
                    var startupManager = new AsyncStartupManager();
                    
                    // 绑定启动进度事件
                    startupManager.ProgressReported += (s, e) =>
                    {
                        try
                        {
                            if (startupForm != null && !startupForm.IsDisposed)
                            {
                                startupForm.UpdateProgress(e.Percentage, e.Message);
                            }
                        }
                        catch (Exception ex)
                        {
                            // 记录进度更新失败，但不中断启动
                            try
                            {
                                Logger.LogWarning($"更新启动进度失败: {ex.Message}");
                            }
                            catch
                            {
                                // 忽略日志记录失败
                            }
                        }
                    };

                    // 绑定启动完成事件
                    startupManager.StartupCompleted += (s, e) =>
                    {
                        try
                        {
                            if (startupForm != null && !startupForm.IsDisposed)
                            {
                                startupForm.CompleteStartup(e.Success, e.Message);
                            }

                            if (e.Success)
                            {
                                // 启动成功，显示主窗体
                                Application.Run(mainForm);
                            }
                            else
                            {
                                // 启动失败，但仍然显示主窗体（降级处理）
                                try
                                {
                                    Logger.LogWarning($"异步启动失败，使用降级模式: {e.Message}");
                                    Application.Run(mainForm);
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogError("降级启动也失败", ex);
                                    Application.Exit();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // 记录启动完成处理失败，但仍然尝试显示主窗体
                            try
                            {
                                Logger.LogError("启动完成事件处理失败", ex);
                                Application.Run(mainForm);
                            }
                            catch (Exception ex2)
                            {
                                Logger.LogError("启动完成处理失败后的降级启动也失败", ex2);
                                Application.Exit();
                            }
                        }
                    };

                    // 开始异步启动
                    try
                    {
                        _ = startupManager.StartAsync();
                    }
                    catch (Exception ex)
                    {
                        // 如果异步启动失败，记录错误但继续显示主窗体
                        try
                        {
                            Logger.LogWarning($"异步启动失败，使用同步模式: {ex.Message}");
                        }
                        catch
                        {
                            // 忽略日志记录失败
                        }
                        
                        // 关闭启动进度窗体
                        if (startupForm != null && !startupForm.IsDisposed)
                        {
                            startupForm.Close();
                        }
                        
                        // 直接显示主窗体
                        Application.Run(mainForm);
                    }
                }
                else
                {
                    // 启动进度窗体创建失败，直接显示主窗体
                    try
                    {
                        Logger.LogInfo("启动进度窗体创建失败，使用同步启动模式");
                    }
                    catch
                    {
                        // 忽略日志记录失败
                    }
                    
                    Application.Run(mainForm);
                }
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
                
                // 确保程序不会静默退出
                Application.Exit();
            }
        }
    }
}