using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYToolsTest
{
    public partial class MainForm : Form
    {
        private object yyToolsInstance;
        private TextBox txtOutput;
        private Button btnCreateCOM;
        private Button btnGetInfo;
        private Button btnDetailedInfo;
        private Button btnInstallMenu;
        private Button btnShowMatch;
        private Button btnRefreshMenu;
        private Button btnClear;
        private Label lblStatus;

        public MainForm()
        {
            InitializeComponent();
            AppendOutput("YYTools 测试程序界面版 v2.1");
            AppendOutput("========================================");
        }

        private void InitializeComponent()
        {
            this.txtOutput = new TextBox();
            this.btnCreateCOM = new Button();
            this.btnGetInfo = new Button();
            this.btnDetailedInfo = new Button();
            this.btnInstallMenu = new Button();
            this.btnShowMatch = new Button();
            this.btnRefreshMenu = new Button();
            this.btnClear = new Button();
            this.lblStatus = new Label();
            this.SuspendLayout();

            // 
            // txtOutput
            // 
            this.txtOutput.Location = new Point(12, 12);
            this.txtOutput.Multiline = true;
            this.txtOutput.ScrollBars = ScrollBars.Vertical;
            this.txtOutput.Size = new Size(560, 300);
            this.txtOutput.ReadOnly = true;
            this.txtOutput.Font = new Font("Consolas", 9F);

            // 
            // btnCreateCOM
            // 
            this.btnCreateCOM.Location = new Point(12, 330);
            this.btnCreateCOM.Size = new Size(100, 30);
            this.btnCreateCOM.Text = "创建COM对象";
            this.btnCreateCOM.UseVisualStyleBackColor = true;
            this.btnCreateCOM.Click += new EventHandler(this.btnCreateCOM_Click);

            // 
            // btnGetInfo
            // 
            this.btnGetInfo.Location = new Point(120, 330);
            this.btnGetInfo.Size = new Size(80, 30);
            this.btnGetInfo.Text = "基本信息";
            this.btnGetInfo.UseVisualStyleBackColor = true;
            this.btnGetInfo.Click += new EventHandler(this.btnGetInfo_Click);

            // 
            // btnDetailedInfo
            // 
            this.btnDetailedInfo.Location = new Point(208, 330);
            this.btnDetailedInfo.Size = new Size(80, 30);
            this.btnDetailedInfo.Text = "详细信息";
            this.btnDetailedInfo.UseVisualStyleBackColor = true;
            this.btnDetailedInfo.Click += new EventHandler(this.btnDetailedInfo_Click);

            // 
            // btnInstallMenu
            // 
            this.btnInstallMenu.Location = new Point(296, 330);
            this.btnInstallMenu.Size = new Size(80, 30);
            this.btnInstallMenu.Text = "安装菜单";
            this.btnInstallMenu.UseVisualStyleBackColor = true;
            this.btnInstallMenu.Click += new EventHandler(this.btnInstallMenu_Click);

            // 
            // btnShowMatch
            // 
            this.btnShowMatch.Location = new Point(12, 370);
            this.btnShowMatch.Size = new Size(100, 30);
            this.btnShowMatch.Text = "显示匹配窗体";
            this.btnShowMatch.UseVisualStyleBackColor = true;
            this.btnShowMatch.Click += new EventHandler(this.btnShowMatch_Click);

            // 
            // btnRefreshMenu
            // 
            this.btnRefreshMenu.Location = new Point(120, 370);
            this.btnRefreshMenu.Size = new Size(80, 30);
            this.btnRefreshMenu.Text = "刷新菜单";
            this.btnRefreshMenu.UseVisualStyleBackColor = true;
            this.btnRefreshMenu.Click += new EventHandler(this.btnRefreshMenu_Click);

            // 
            // btnClear
            // 
            this.btnClear.Location = new Point(492, 330);
            this.btnClear.Size = new Size(80, 30);
            this.btnClear.Text = "清空输出";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new EventHandler(this.btnClear_Click);

            // 
            // lblStatus
            // 
            this.lblStatus.Location = new Point(12, 415);
            this.lblStatus.Size = new Size(560, 23);
            this.lblStatus.Text = "状态: 请先点击'创建COM对象'";
            this.lblStatus.BackColor = Color.LightYellow;
            this.lblStatus.TextAlign = ContentAlignment.MiddleLeft;

            // 
            // MainForm
            // 
            this.ClientSize = new Size(584, 450);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.btnCreateCOM);
            this.Controls.Add(this.btnGetInfo);
            this.Controls.Add(this.btnDetailedInfo);
            this.Controls.Add(this.btnInstallMenu);
            this.Controls.Add(this.btnShowMatch);
            this.Controls.Add(this.btnRefreshMenu);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.lblStatus);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "YYTools 测试程序 v2.1";
            this.ResumeLayout(false);

            // 初始状态下除了创建COM对象按钮，其他都禁用
            SetButtonsEnabled(false);
            btnCreateCOM.Enabled = true;
            btnClear.Enabled = true;
        }

        private void SetButtonsEnabled(bool enabled)
        {
            btnGetInfo.Enabled = enabled;
            btnDetailedInfo.Enabled = enabled;
            btnInstallMenu.Enabled = enabled;
            btnShowMatch.Enabled = enabled;
            btnRefreshMenu.Enabled = enabled;
        }

        private void AppendOutput(string text)
        {
            if (txtOutput.InvokeRequired)
            {
                txtOutput.Invoke(new Action<string>(AppendOutput), text);
            }
            else
            {
                txtOutput.AppendText(text + Environment.NewLine);
                txtOutput.SelectionStart = txtOutput.Text.Length;
                txtOutput.ScrollToCaret();
            }
        }

        private void UpdateStatus(string status, bool isError = false)
        {
            lblStatus.Text = "状态: " + status;
            lblStatus.BackColor = isError ? Color.LightPink : Color.LightGreen;
        }

        private void btnCreateCOM_Click(object sender, EventArgs e)
        {
            try
            {
                AppendOutput("正在创建COM对象...");
                yyToolsInstance = Activator.CreateInstance(Type.GetTypeFromProgID("YYTools.ExcelAddin"));
                
                if (yyToolsInstance != null)
                {
                    AppendOutput("✓ COM对象创建成功");
                    UpdateStatus("COM对象已创建，可以使用其他功能");
                    SetButtonsEnabled(true);
                    btnCreateCOM.Text = "重新创建";
                }
                else
                {
                    AppendOutput("✗ COM对象创建失败");
                    UpdateStatus("COM对象创建失败", true);
                }
            }
            catch (Exception ex)
            {
                AppendOutput("✗ 创建COM对象异常: " + ex.Message);
                UpdateStatus("COM对象创建异常: " + ex.Message, true);
                AppendOutput("可能的原因:");
                AppendOutput("1. YYTools.dll未正确注册");
                AppendOutput("2. 需要管理员权限");
                AppendOutput("3. .NET Framework版本不兼容");
            }
        }

        private void btnGetInfo_Click(object sender, EventArgs e)
        {
            if (yyToolsInstance == null)
            {
                AppendOutput("✗ 请先创建COM对象");
                return;
            }

            try
            {
                AppendOutput("正在获取基本应用程序信息...");
                string info = (string)yyToolsInstance.GetType().InvokeMember("GetApplicationInfo",
                    System.Reflection.BindingFlags.InvokeMethod, null, yyToolsInstance, null);
                AppendOutput("✓ 基本信息: " + info);
                UpdateStatus("基本信息获取成功");
            }
            catch (Exception ex)
            {
                AppendOutput("✗ 获取基本信息失败: " + ex.Message);
                UpdateStatus("获取基本信息失败", true);
            }
        }

        private void btnDetailedInfo_Click(object sender, EventArgs e)
        {
            if (yyToolsInstance == null)
            {
                AppendOutput("✗ 请先创建COM对象");
                return;
            }

            try
            {
                AppendOutput("正在获取详细应用程序信息...");
                string info = (string)yyToolsInstance.GetType().InvokeMember("GetDetailedApplicationInfo",
                    System.Reflection.BindingFlags.InvokeMethod, null, yyToolsInstance, null);
                AppendOutput("✓ 详细信息:");
                AppendOutput(info);
                UpdateStatus("详细信息获取成功");
            }
            catch (Exception ex)
            {
                AppendOutput("✗ 获取详细信息失败: " + ex.Message);
                UpdateStatus("获取详细信息失败", true);
            }
        }

        private void btnInstallMenu_Click(object sender, EventArgs e)
        {
            if (yyToolsInstance == null)
            {
                AppendOutput("✗ 请先创建COM对象");
                return;
            }

            try
            {
                AppendOutput("正在安装菜单...");
                string result = (string)yyToolsInstance.GetType().InvokeMember("InstallMenu",
                    System.Reflection.BindingFlags.InvokeMethod, null, yyToolsInstance, null);
                AppendOutput("✓ 菜单安装结果: " + result);
                UpdateStatus("菜单安装操作完成，请检查WPS/Excel工具栏");
                
                // 显示额外信息
                MessageBox.Show("菜单安装操作完成！\n\n" + result + "\n\n请检查WPS表格或Excel的工具栏是否出现'YY工具'菜单。\n\n如果没有出现，可能是:\n1. WPS/Excel权限限制\n2. 需要重启WPS/Excel\n3. CommandBars API不支持", 
                    "菜单安装结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendOutput("✗ 安装菜单失败: " + ex.Message);
                UpdateStatus("安装菜单失败", true);
                MessageBox.Show("安装菜单失败:\n\n" + ex.Message + "\n\n可能的原因:\n1. WPS/Excel未启动\n2. 工作簿未打开\n3. CommandBars API访问受限", 
                    "菜单安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnShowMatch_Click(object sender, EventArgs e)
        {
            if (yyToolsInstance == null)
            {
                AppendOutput("✗ 请先创建COM对象");
                return;
            }

            try
            {
                AppendOutput("正在显示匹配窗体...");
                yyToolsInstance.GetType().InvokeMember("ShowMatchForm",
                    System.Reflection.BindingFlags.InvokeMethod, null, yyToolsInstance, null);
                AppendOutput("✓ 匹配窗体显示成功");
                UpdateStatus("匹配窗体显示成功");
            }
            catch (Exception ex)
            {
                AppendOutput("✗ 显示匹配窗体失败: " + ex.Message);
                UpdateStatus("显示匹配窗体失败", true);
            }
        }

        private void btnRefreshMenu_Click(object sender, EventArgs e)
        {
            if (yyToolsInstance == null)
            {
                AppendOutput("✗ 请先创建COM对象");
                return;
            }

            try
            {
                AppendOutput("正在刷新菜单...");
                yyToolsInstance.GetType().InvokeMember("RefreshMenu",
                    System.Reflection.BindingFlags.InvokeMethod, null, yyToolsInstance, null);
                AppendOutput("✓ 菜单刷新成功");
                UpdateStatus("菜单刷新成功");
            }
            catch (Exception ex)
            {
                AppendOutput("✗ 刷新菜单失败: " + ex.Message);
                UpdateStatus("刷新菜单失败", true);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
            AppendOutput("YYTools 测试程序界面版 v2.1");
            AppendOutput("========================================");
            UpdateStatus("输出已清空");
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (yyToolsInstance != null)
            {
                yyToolsInstance = null;
            }
            base.OnFormClosing(e);
        }
    }
} 