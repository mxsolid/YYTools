using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 用户引导窗体
    /// </summary>
    public partial class UserGuideForm : Form
    {
        private TabControl tabControl;
        private TabPage quickStartTab;
        private TabPage detailedGuideTab;
        private TabPage faqTab;
        private TabPage aboutTab;

        public UserGuideForm()
        {
            InitializeComponent();
            InitializeGuideContent();
            ApplyUIEnhancement();
        }

        private void InitializeComponent()
        {
            this.Text = "YY工具使用指南";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
        }

        private void InitializeGuideContent()
        {
            // 创建主控件
            tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;

            // 快速开始标签页
            CreateQuickStartTab();

            // 详细指南标签页
            CreateDetailedGuideTab();

            // 常见问题标签页
            CreateFaqTab();

            // 关于标签页
            CreateAboutTab();

            // 添加标签页到控件
            tabControl.TabPages.Add(quickStartTab);
            tabControl.TabPages.Add(detailedGuideTab);
            tabControl.TabPages.Add(faqTab);
            tabControl.TabPages.Add(aboutTab);

            // 添加关闭按钮
            var closeButton = new Button
            {
                Text = "关闭",
                Size = new Size(80, 30),
                Location = new Point(this.Width - 100, this.Height - 50),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            closeButton.Click += (s, e) => this.Close();

            // 添加控件到窗体
            this.Controls.Add(tabControl);
            this.Controls.Add(closeButton);
        }

        private void CreateQuickStartTab()
        {
            quickStartTab = new TabPage("快速开始");
            quickStartTab.BackColor = Color.White;

            var richTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("微软雅黑", 10F)
            };

            richTextBox.Text = Constants.QuickStartGuide;

            quickStartTab.Controls.Add(richTextBox);
        }

        private void CreateDetailedGuideTab()
        {
            detailedGuideTab = new TabPage("详细指南");
            detailedGuideTab.BackColor = Color.White;

            var richTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("微软雅黑", 10F)
            };

            richTextBox.Text = Constants.DetailedGuide;

            detailedGuideTab.Controls.Add(richTextBox);
        }

        private void CreateFaqTab()
        {
            faqTab = new TabPage("常见问题");
            faqTab.BackColor = Color.White;

            var richTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("微软雅黑", 10F)
            };

            richTextBox.Text = Constants.FrequentlyAskedQuestions;

            faqTab.Controls.Add(richTextBox);
        }

        private void CreateAboutTab()
        {
            aboutTab = new TabPage("关于");
            aboutTab.BackColor = Color.White;

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            // 标题
            var titleLabel = new Label
            {
                Text = Constants.AppName,
                Font = new Font("微软雅黑", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204),
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 40),
                Location = new Point(200, 50)
            };

            // 版本信息
            var versionLabel = new Label
            {
                Text = $"版本: {Constants.AppVersion}",
                Font = new Font("微软雅黑", 12F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 30),
                Location = new Point(200, 100)
            };

            // 公司信息
            var companyLabel = new Label
            {
                Text = $"公司: {Constants.AppCompany}",
                Font = new Font("微软雅黑", 12F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 30),
                Location = new Point(200, 140)
            };

            // 功能特性
            var featuresLabel = new Label
            {
                Text = Constants.AppFeatures,
                Font = new Font("微软雅黑", 10F),
                ForeColor = Color.Black,
                Size = new Size(400, 200),
                Location = new Point(200, 200)
            };

            // 版权信息
            var copyrightLabel = new Label
            {
                Text = "© 2024 YY Tools. 保留所有权利。",
                Font = new Font("微软雅黑", 9F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 30),
                Location = new Point(200, 420)
            };

            // 添加控件
            panel.Controls.Add(titleLabel);
            panel.Controls.Add(versionLabel);
            panel.Controls.Add(companyLabel);
            panel.Controls.Add(featuresLabel);
            panel.Controls.Add(copyrightLabel);

            aboutTab.Controls.Add(panel);
        }

        private void ApplyUIEnhancement()
        {
            try
            {
                // 应用UI美化
                UIEnhancer.EnhanceForm(this);
                UIEnhancer.EnhanceAllControls(this);

                // 设置标签页样式
                tabControl.Font = new Font("微软雅黑", 9F);
                tabControl.BackColor = Color.White;
            }
            catch (Exception ex)
            {
                Logger.LogError("应用UI美化失败", ex);
            }
        }

        /// <summary>
        /// 显示用户引导窗体
        /// </summary>
        public static void ShowUserGuide(IWin32Window owner = null)
        {
            try
            {
                var form = new UserGuideForm();
                form.ShowDialog(owner);
            }
            catch (Exception ex)
            {
                Logger.LogError("显示用户引导窗体失败", ex);
                MessageBox.Show("显示用户引导失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}