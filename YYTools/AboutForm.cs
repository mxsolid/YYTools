using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 关于窗体 - 显示程序信息和版本详情
    /// 采用现代化设计风格，集成版本管理器
    /// </summary>
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            InitializeAboutForm();
            LoadVersionInfo();
        }

        /// <summary>
        /// 初始化关于窗体
        /// </summary>
        private void InitializeAboutForm()
        {
            try
            {
                // 设置窗体属性
                this.Text = "关于 YY运单匹配工具";
                this.Size = new Size(500, 600);
                this.StartPosition = FormStartPosition.CenterParent;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = false;
                this.BackColor = Color.FromArgb(248, 248, 248);

                // 设置DPI感知
                this.AutoScaleMode = AutoScaleMode.Dpi;
                this.AutoScaleDimensions = new SizeF(96F, 96F);

                // 创建控件
                CreateControls();

                Logger.LogInfo("关于窗体初始化完成");
            }
            catch (Exception ex)
            {
                Logger.LogError("初始化关于窗体失败", ex);
            }
        }

        /// <summary>
        /// 创建控件
        /// </summary>
        private void CreateControls()
        {
            try
            {
                // 标题标签
                var lblTitle = new Label
                {
                    Text = "YY运单匹配工具",
                    Font = new Font("微软雅黑", 18, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 122, 204),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(460, 40),
                    Location = new Point(20, 20)
                };
                this.Controls.Add(lblTitle);

                // 版本标签
                var lblVersion = new Label
                {
                    Text = "版本 " + VersionManager.DisplayVersion,
                    Font = new Font("微软雅黑", 12, FontStyle.Regular),
                    ForeColor = Color.FromArgb(100, 100, 100),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(460, 25),
                    Location = new Point(20, 70)
                };
                this.Controls.Add(lblVersion);

                // 构建信息标签
                var lblBuild = new Label
                {
                    Text = "构建 " + VersionManager.BuildVersion,
                    Font = new Font("微软雅黑", 10, FontStyle.Regular),
                    ForeColor = Color.FromArgb(120, 120, 120),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(460, 20),
                    Location = new Point(20, 100)
                };
                this.Controls.Add(lblBuild);

                // 分隔线
                var separator1 = new Panel
                {
                    BackColor = Color.FromArgb(200, 200, 200),
                    Size = new Size(460, 1),
                    Location = new Point(20, 130)
                };
                this.Controls.Add(separator1);

                // 描述标签
                var lblDescription = new Label
                {
                    Text = "专业的运单匹配工具，支持Excel和WPS，具备智能列选择和高性能处理能力。",
                    Font = new Font("微软雅黑", 10, FontStyle.Regular),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Size = new Size(460, 40),
                    Location = new Point(20, 150)
                };
                this.Controls.Add(lblDescription);

                // 特性列表
                var lblFeatures = new Label
                {
                    Text = "主要特性：",
                    Font = new Font("微软雅黑", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    Size = new Size(460, 20),
                    Location = new Point(20, 200)
                };
                this.Controls.Add(lblFeatures);

                var lblFeature1 = new Label
                {
                    Text = "• 智能DPI适配，完美支持高分辨率显示器",
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Size = new Size(460, 18),
                    Location = new Point(30, 225)
                };
                // this.Controls.Add(lblFeature1);

                var lblFeature2 = new Label
                {
                    Text = "• 多显示器支持，自动适配不同DPI设置",
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Size = new Size(460, 18),
                    Location = new Point(30, 245)
                };
                // this.Controls.Add(lblFeature2);

                var lblFeature3 = new Label
                {
                    Text = "• Excel/WPS双重兼容，无需额外配置",
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Size = new Size(460, 18),
                    Location = new Point(30, 265)
                };
                this.Controls.Add(lblFeature3);

                var lblFeature4 = new Label
                {
                    Text = "• 高性能处理引擎，支持大数据量操作",
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Size = new Size(460, 18),
                    Location = new Point(30, 285)
                };
                this.Controls.Add(lblFeature4);

                // 分隔线
                var separator2 = new Panel
                {
                    BackColor = Color.FromArgb(200, 200, 200),
                    Size = new Size(460, 1),
                    Location = new Point(20, 320)
                };
                this.Controls.Add(separator2);

                // 作者信息
                var lblAuthor = new Label
                {
                    Text = "作者信息：",
                    Font = new Font("微软雅黑", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    Size = new Size(460, 20),
                    Location = new Point(20, 340)
                };
                this.Controls.Add(lblAuthor);

                var lblAuthorName = new Label
                {
                    Text = "皮皮熊",
                    Font = new Font("微软雅黑", 10, FontStyle.Regular),
                    ForeColor = Color.FromArgb(0, 122, 204),
                    Size = new Size(460, 20),
                    Location = new Point(30, 365)
                };
                this.Controls.Add(lblAuthorName);

                var lblEmail = new Label
                {
                    Text = "联系邮箱：oyxo@qq.com",
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Size = new Size(460, 18),
                    Location = new Point(30, 385)
                };
                this.Controls.Add(lblEmail);

                // 分隔线
                var separator3 = new Panel
                {
                    BackColor = Color.FromArgb(200, 200, 200),
                    Size = new Size(460, 1),
                    Location = new Point(20, 420)
                };
                this.Controls.Add(separator3);

                // 技术信息
                var lblTech = new Label
                {
                    Text = "技术信息：",
                    Font = new Font("微软雅黑", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    Size = new Size(460, 20),
                    Location = new Point(20, 440)
                };
                this.Controls.Add(lblTech);

                var lblFramework = new Label
                {
                    Text = "框架版本：.NET Framework " + Environment.Version,
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Size = new Size(460, 18),
                    Location = new Point(30, 465)
                };
                this.Controls.Add(lblFramework);

                var lblOS = new Label
                {
                    Text = "操作系统：" + Environment.OSVersion,
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Size = new Size(460, 18),
                    Location = new Point(30, 485)
                };
                this.Controls.Add(lblOS);

                // 确定按钮
                var btnOK = new Button
                {
                    Text = "确定",
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    Size = new Size(80, 32),
                    Location = new Point(210, 520),
                    BackColor = Color.FromArgb(0, 122, 204),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    UseVisualStyleBackColor = false
                };
                btnOK.Click += (s, e) => this.Close();
                btnOK.FlatAppearance.BorderSize = 0;
                this.Controls.Add(btnOK);

                // 查看详细信息按钮
                var btnDetails = new Button
                {
                    Text = "详细信息",
                    Font = new Font("微软雅黑", 9, FontStyle.Regular),
                    Size = new Size(80, 32),
                    Location = new Point(300, 520),
                    BackColor = Color.FromArgb(240, 240, 240),
                    ForeColor = Color.FromArgb(60, 60, 60),
                    FlatStyle = FlatStyle.Flat,
                    UseVisualStyleBackColor = false
                };
                btnDetails.Click += (s, e) => ShowDetailedInfo();
                btnDetails.FlatAppearance.BorderSize = 1;
                btnDetails.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
                this.Controls.Add(btnDetails);

                Logger.LogInfo("关于窗体控件创建完成");
            }
            catch (Exception ex)
            {
                Logger.LogError("创建关于窗体控件失败", ex);
            }
        }

        /// <summary>
        /// 加载版本信息
        /// </summary>
        private void LoadVersionInfo()
        {
            try
            {
                // 版本信息已在控件创建时加载
                Logger.LogInfo("版本信息加载完成");
            }
            catch (Exception ex)
            {
                Logger.LogError("加载版本信息失败", ex);
            }
        }

        /// <summary>
        /// 显示详细信息
        /// </summary>
        private void ShowDetailedInfo()
        {
            try
            {
                var detailedInfo = VersionManager.GetFullVersionInfo();
                
                var detailForm = new Form
                {
                    Text = "详细信息",
                    Size = new Size(600, 500),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    ShowInTaskbar = false,
                    BackColor = Color.FromArgb(248, 248, 248)
                };

                var textBox = new TextBox
                {
                    Text = detailedInfo,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Font = new Font("Consolas", 9),
                    Size = new Size(560, 400),
                    Location = new Point(20, 20),
                    BackColor = Color.White
                };

                var btnClose = new Button
                {
                    Text = "关闭",
                    Size = new Size(80, 32),
                    Location = new Point(250, 430),
                    BackColor = Color.FromArgb(0, 122, 204),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    UseVisualStyleBackColor = false
                };
                btnClose.Click += (s, e) => detailForm.Close();
                btnClose.FlatAppearance.BorderSize = 0;

                detailForm.Controls.Add(textBox);
                detailForm.Controls.Add(btnClose);

                detailForm.ShowDialog(this);
                Logger.LogInfo("详细信息窗体显示完成");
            }
            catch (Exception ex)
            {
                Logger.LogError("显示详细信息失败", ex);
                MessageBox.Show("显示详细信息失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
