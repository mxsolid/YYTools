using System;
using System.Drawing;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace YYTools
{
    /// <summary>
    /// 启动进度窗体
    /// </summary>
    public partial class StartupProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;
        private Panel mainPanel;
        private bool _isClosing = false;

        public StartupProgressForm()
        {
            InitializeComponent();
            ApplyUIEnhancement();
        }

        private void InitializeComponent()
        {
            this.Text = "YY工具 - 正在启动";
            this.Size = new Size(500, 200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.TopMost = true;

            // 主面板
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(20)
            };

            // 标题标签
            var lblTitle = new Label
            {
                Text = "YY 运单匹配工具",
                Font = new Font("微软雅黑", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204),
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(460, 40),
                Location = new Point(20, 20)
            };

            // 状态标签
            lblStatus = new Label
            {
                Text = "正在初始化应用程序...",
                Font = new Font("微软雅黑", 10F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleLeft,
                Size = new Size(460, 25),
                Location = new Point(20, 70)
            };

            // 进度条
            progressBar = new ProgressBar
            {
                Size = new Size(460, 25),
                Location = new Point(20, 100),
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };

            // 进度标签
            lblProgress = new Label
            {
                Text = "0%",
                Font = new Font("微软雅黑", 9F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleRight,
                Size = new Size(460, 20),
                Location = new Point(20, 130)
            };

            // 添加控件到主面板
            mainPanel.Controls.AddRange(new Control[] 
            {
                lblTitle, lblStatus, progressBar, lblProgress
            });

            // 添加主面板到窗体
            this.Controls.Add(mainPanel);

            // 设置窗体样式
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 更新进度
        /// </summary>
        public void UpdateProgress(int percentage, string status)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<int, string>(UpdateProgress), percentage, status);
                    return;
                }

                if (_isClosing) return;

                // 更新进度条
                progressBar.Value = Math.Max(0, Math.Min(100, percentage));
                
                // 更新状态文本
                lblStatus.Text = status ?? string.Empty;
                
                // 更新进度百分比
                lblProgress.Text = $"{percentage}%";

                // 刷新界面
                this.Refresh();
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                // 记录错误但不抛出异常
                try
                {
                    Logger.LogError($"更新启动进度失败: {ex.Message}", ex);
                }
                catch
                {
                    // 忽略日志记录失败
                }
            }
        }

        /// <summary>
        /// 完成启动
        /// </summary>
        public void CompleteStartup(bool success, string message)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<bool, string>(CompleteStartup), success, message);
                    return;
                }

                if (_isClosing) return;

                if (success)
                {
                    // 启动成功
                    progressBar.Value = 100;
                    lblStatus.Text = "启动完成！";
                    lblProgress.Text = "100%";
                    
                    // 延迟关闭窗体
                    Timer closeTimer = new Timer();
                    closeTimer.Interval = 1000; // 1秒后关闭
                    closeTimer.Tick += (s, e) =>
                    {
                        closeTimer.Stop();
                        closeTimer.Dispose();
                        this.Close();
                    };
                    closeTimer.Start();
                }
                else
                {
                    // 启动失败
                    lblStatus.Text = $"启动失败: {message}";
                    lblStatus.ForeColor = Color.Red;
                    
                    // 显示错误信息
                    MessageBox.Show($"启动失败：{message}\n\n请检查Excel应用程序是否正常运行。", 
                        "启动错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                try
                {
                    Logger.LogError($"完成启动处理失败: {ex.Message}", ex);
                }
                catch
                {
                    // 忽略日志记录失败
                }
                
                this.Close();
            }
        }

        /// <summary>
        /// 应用UI美化
        /// </summary>
        private void ApplyUIEnhancement()
        {
            try
            {
                // 应用基础美化
                this.BackColor = Color.White;
                this.Font = new Font("微软雅黑", 9F, FontStyle.Regular);

                // 美化进度条
                progressBar.ForeColor = Color.FromArgb(0, 122, 204);
                progressBar.BackColor = Color.FromArgb(240, 240, 240);

                // 设置标签样式
                lblStatus.Font = new Font("微软雅黑", 10F, FontStyle.Regular);
                lblProgress.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
            }
            catch (Exception ex)
            {
                try
                {
                    Logger.LogError($"应用UI美化失败: {ex.Message}", ex);
                }
                catch
                {
                    // 忽略日志记录失败
                }
            }
        }

        /// <summary>
        /// 窗体关闭事件
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _isClosing = true;
            base.OnFormClosing(e);
        }

        /// <summary>
        /// 显示启动进度窗体
        /// </summary>
        public static StartupProgressForm ShowStartupProgress()
        {
            try
            {
                var form = new StartupProgressForm();
                
                // 确保窗体正确显示
                form.Show();
                form.BringToFront();
                form.Refresh();
                
                // 强制处理消息队列
                Application.DoEvents();
                
                return form;
            }
            catch (Exception ex)
            {
                // 记录错误但不抛出异常
                try
                {
                    Logger.LogError("显示启动进度窗体失败", ex);
                }
                catch
                {
                    // 忽略日志记录失败
                }
                
                // 返回null，让调用者使用降级模式
                return null;
            }
        }
    }
}