using System;
using System.Drawing;
using System.Windows.Forms;

namespace DpiTest
{
    public partial class DpiTestForm : Form
    {
        private Label lblDpiInfo;
        private Button btnTestDpi;
        private TextBox txtTest;
        private ComboBox cmbTest;
        private GroupBox gbTest;

        public DpiTestForm()
        {
            InitializeComponent();
            FixDpiIssues();
        }

        private void InitializeComponent()
        {
            this.lblDpiInfo = new Label();
            this.btnTestDpi = new Button();
            this.txtTest = new TextBox();
            this.cmbTest = new ComboBox();
            this.gbTest = new GroupBox();
            this.SuspendLayout();

            // lblDpiInfo
            this.lblDpiInfo.AutoSize = true;
            this.lblDpiInfo.Location = new Point(12, 9);
            this.lblDpiInfo.Name = "lblDpiInfo";
            this.lblDpiInfo.Size = new Size(100, 23);
            this.lblDpiInfo.Text = "DPI信息:";

            // btnTestDpi
            this.btnTestDpi.Location = new Point(12, 40);
            this.btnTestDpi.Name = "btnTestDpi";
            this.btnTestDpi.Size = new Size(100, 30);
            this.btnTestDpi.Text = "测试DPI";
            this.btnTestDpi.Click += new EventHandler(this.btnTestDpi_Click);

            // txtTest
            this.txtTest.Location = new Point(12, 80);
            this.txtTest.Name = "txtTest";
            this.txtTest.Size = new Size(200, 25);
            this.txtTest.Text = "测试文本框";

            // cmbTest
            this.cmbTest.Location = new Point(12, 120);
            this.cmbTest.Name = "cmbTest";
            this.cmbTest.Size = new Size(200, 25);
            this.cmbTest.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbTest.Items.AddRange(new object[] { "测试项目1", "测试项目2", "测试项目3" });
            this.cmbTest.SelectedIndex = 0;

            // gbTest
            this.gbTest.Location = new Point(12, 160);
            this.gbTest.Name = "gbTest";
            this.gbTest.Size = new Size(300, 100);
            this.gbTest.Text = "测试分组框";

            // DpiTestForm
            this.AutoScaleDimensions = new SizeF(6F, 12F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(400, 300);
            this.Controls.Add(this.lblDpiInfo);
            this.Controls.Add(this.btnTestDpi);
            this.Controls.Add(this.txtTest);
            this.Controls.Add(this.cmbTest);
            this.Controls.Add(this.gbTest);
            this.Name = "DpiTestForm";
            this.Text = "DPI测试程序";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void btnTestDpi_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取当前DPI信息
                float dpiScale = GetCurrentDpiScale();
                string dpiInfo = $"DPI缩放: {dpiScale:F2}\n" +
                               $"窗体字体: {this.Font.Size:F1}\n" +
                               $"标签字体: {lblDpiInfo.Font.Size:F1}\n" +
                               $"按钮字体: {btnTestDpi.Font.Size:F1}\n" +
                               $"文本框字体: {txtTest.Font.Size:F1}\n" +
                               $"下拉框字体: {cmbTest.Font.Size:F1}\n" +
                               $"分组框字体: {gbTest.Font.Size:F1}";

                MessageBox.Show(dpiInfo, "DPI测试结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"测试DPI时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 修复DPI问题
        /// </summary>
        private void FixDpiIssues()
        {
            try
            {
                // 获取当前DPI信息
                float dpiScale = GetCurrentDpiScale();
                
                // 限制缩放比例，防止过大
                float maxScale = 1.25f;
                float actualScale = Math.Min(dpiScale, maxScale);
                
                // 修复字体大小
                FixFontSizes(actualScale);
                
                // 修复控件尺寸
                FixControlDimensions(actualScale);
                
                // 修复窗体尺寸
                FixFormDimensions(actualScale);
                
                // 更新DPI信息标签
                UpdateDpiInfoLabel(actualScale);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"修复DPI问题时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取当前DPI缩放比例
        /// </summary>
        private float GetCurrentDpiScale()
        {
            try
            {
                using (Graphics g = this.CreateGraphics())
                {
                    return g.DpiX / 96.0f;
                }
            }
            catch
            {
                return 1.0f;
            }
        }

        /// <summary>
        /// 修复字体大小
        /// </summary>
        private void FixFontSizes(float scale)
        {
            try
            {
                // 修复窗体字体
                if (this.Font != null)
                {
                    float newSize = Math.Max(8.0f, Math.Min(12.0f, this.Font.Size * scale));
                    this.Font = new Font(this.Font.FontFamily, newSize, this.Font.Style);
                }

                // 修复标签字体
                if (lblDpiInfo.Font != null)
                {
                    float newSize = Math.Max(8.0f, Math.Min(12.0f, lblDpiInfo.Font.Size * scale));
                    lblDpiInfo.Font = new Font(lblDpiInfo.Font.FontFamily, newSize, lblDpiInfo.Font.Style);
                }

                // 修复按钮字体
                if (btnTestDpi.Font != null)
                {
                    float newSize = Math.Max(8.0f, Math.Min(12.0f, btnTestDpi.Font.Size * scale));
                    btnTestDpi.Font = new Font(btnTestDpi.Font.FontFamily, newSize, btnTestDpi.Font.Style);
                }

                // 修复文本框字体
                if (txtTest.Font != null)
                {
                    float newSize = Math.Max(8.0f, Math.Min(12.0f, txtTest.Font.Size * scale));
                    txtTest.Font = new Font(txtTest.Font.FontFamily, newSize, txtTest.Font.Style);
                }

                // 修复下拉框字体
                if (cmbTest.Font != null)
                {
                    float newSize = Math.Max(8.0f, Math.Min(12.0f, cmbTest.Font.Size * scale));
                    cmbTest.Font = new Font(cmbTest.Font.FontFamily, newSize, cmbTest.Font.Style);
                }

                // 修复分组框字体
                if (gbTest.Font != null)
                {
                    float newSize = Math.Max(8.0f, Math.Min(12.0f, gbTest.Font.Size * scale));
                    gbTest.Font = new Font(gbTest.Font.FontFamily, newSize, gbTest.Font.Style);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"修复字体大小失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 修复控件尺寸
        /// </summary>
        private void FixControlDimensions(float scale)
        {
            try
            {
                // 限制最大缩放
                float maxScale = 1.2f;
                float actualScale = Math.Min(scale, maxScale);

                // 修复按钮尺寸
                int newHeight = (int)(btnTestDpi.Height * actualScale);
                newHeight = Math.Max(25, Math.Min(newHeight, 45));
                btnTestDpi.Height = newHeight;

                // 修复文本框尺寸
                newHeight = (int)(txtTest.Height * actualScale);
                newHeight = Math.Max(20, Math.Min(newHeight, 40));
                txtTest.Height = newHeight;

                // 修复下拉框尺寸
                newHeight = (int)(cmbTest.Height * actualScale);
                newHeight = Math.Max(20, Math.Min(newHeight, 40));
                cmbTest.Height = newHeight;
                cmbTest.DropDownWidth = Math.Max(cmbTest.Width + 50, 300);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"修复控件尺寸失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 修复窗体尺寸
        /// </summary>
        private void FixFormDimensions(float scale)
        {
            try
            {
                // 限制最大缩放
                float maxScale = 1.2f;
                float actualScale = Math.Min(scale, maxScale);

                // 调整窗体大小
                if (this.Size.Width > 0 && this.Size.Height > 0)
                {
                    Size newSize = new Size(
                        (int)(this.Size.Width * actualScale),
                        (int)(this.Size.Height * actualScale)
                    );

                    // 确保窗体不会超出屏幕边界
                    Rectangle screenBounds = Screen.GetWorkingArea(this);
                    if (newSize.Width > screenBounds.Width)
                    {
                        newSize.Width = screenBounds.Width - 100;
                    }
                    if (newSize.Height > screenBounds.Height)
                    {
                        newSize.Height = screenBounds.Height - 100;
                    }

                    // 限制最大尺寸
                    newSize.Width = Math.Min(newSize.Width, 900);
                    newSize.Height = Math.Min(newSize.Height, 700);

                    this.Size = newSize;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"修复窗体尺寸失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新DPI信息标签
        /// </summary>
        private void UpdateDpiInfoLabel(float scale)
        {
            try
            {
                lblDpiInfo.Text = $"DPI缩放: {scale:F2}";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新DPI信息标签失败: {ex.Message}");
            }
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new DpiTestForm());
        }
    }
}