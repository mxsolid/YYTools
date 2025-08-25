using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 列选择对话框
    /// </summary>
    public partial class ColumnSelectionForm : Form
    {
        private TextBox txtColumn;
        private Button btnSelectFromSheet;
        private Button btnOK;
        private Button btnCancel;
        private Label lblInstruction;

        public string SelectedColumn { get; private set; }

        public ColumnSelectionForm(string currentColumn = "")
        {
            InitializeComponent();
            txtColumn.Text = currentColumn;
            SelectedColumn = currentColumn;
            
            // 简化的聚焦设置
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = true;
            this.TopMost = true;
            
            // 窗体显示后设置焦点
            this.Shown += (s, e) => 
            {
                this.TopMost = false;
                txtColumn.Focus();
                txtColumn.SelectAll();
            };
        }

        private void InitializeComponent()
        {
            this.txtColumn = new TextBox();
            this.btnSelectFromSheet = new Button();
            this.btnOK = new Button();
            this.btnCancel = new Button();
            this.lblInstruction = new Label();
            this.SuspendLayout();

            // 
            // lblInstruction
            // 
            this.lblInstruction.Location = new Point(12, 15);
            this.lblInstruction.Size = new Size(360, 40);
            this.lblInstruction.Text = "请输入列号（如A、B、AA等）或点击从表格选择按钮：";
            this.lblInstruction.TextAlign = ContentAlignment.MiddleLeft;

            // 
            // txtColumn
            // 
            this.txtColumn.Location = new Point(12, 60);
            this.txtColumn.Size = new Size(100, 23);
            this.txtColumn.Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Regular);
            this.txtColumn.TextAlign = HorizontalAlignment.Center;
            this.txtColumn.TextChanged += new EventHandler(this.txtColumn_TextChanged);

            // 
            // btnSelectFromSheet
            // 
            this.btnSelectFromSheet.Location = new Point(130, 58);
            this.btnSelectFromSheet.Size = new Size(120, 27);
            this.btnSelectFromSheet.Text = "从表格选择";
            this.btnSelectFromSheet.UseVisualStyleBackColor = true;
            this.btnSelectFromSheet.Click += new EventHandler(this.btnSelectFromSheet_Click);

            // 
            // btnOK
            // 
            this.btnOK.Location = new Point(200, 100);
            this.btnOK.Size = new Size(75, 30);
            this.btnOK.Text = "确定";
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new EventHandler(this.btnOK_Click);

            // 
            // btnCancel
            // 
            this.btnCancel.Location = new Point(285, 100);
            this.btnCancel.Size = new Size(75, 30);
            this.btnCancel.Text = "取消";
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.UseVisualStyleBackColor = true;

            // 
            // ColumnSelectionForm
            // 
            this.ClientSize = new Size(380, 150);
            this.Controls.Add(this.lblInstruction);
            this.Controls.Add(this.txtColumn);
            this.Controls.Add(this.btnSelectFromSheet);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "选择列";
            this.AcceptButton = this.btnOK;
            this.CancelButton = this.btnCancel;
            this.ResumeLayout(false);
        }

        private void txtColumn_TextChanged(object sender, EventArgs e)
        {
            // 实时验证输入的列号
            string input = txtColumn.Text.Trim().ToUpper();
            btnOK.Enabled = string.IsNullOrEmpty(input) || ExcelHelper.IsValidColumnLetter(input);
            
            if (btnOK.Enabled && !string.IsNullOrEmpty(input))
            {
                txtColumn.BackColor = Color.White;
            }
            else if (!string.IsNullOrEmpty(input))
            {
                txtColumn.BackColor = Color.LightPink;
            }
            else
            {
                txtColumn.BackColor = Color.White;
            }
        }

        private void btnSelectFromSheet_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取WPS/Excel应用程序（使用动态COM调用）
                object app = YYTools.ExcelAddin.GetExcelApplication();
                if (app == null)
                {
                    MessageBox.Show("无法连接到WPS表格或Excel！", "错误", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 临时隐藏对话框
                this.Visible = false;

                try
                {
                    // 激活WPS/Excel窗口
                    app.GetType().InvokeMember("Visible", 
                        System.Reflection.BindingFlags.SetProperty, null, app, new object[] { true });
                    
                    object activeWindow = app.GetType().InvokeMember("ActiveWindow", 
                        System.Reflection.BindingFlags.GetProperty, null, app, null);
                    if (activeWindow != null)
                    {
                        activeWindow.GetType().InvokeMember("Activate", 
                            System.Reflection.BindingFlags.InvokeMethod, null, activeWindow, null);
                    }

                    // 提示用户选择
                    DialogResult result = MessageBox.Show(
                        "请在WPS表格或Excel中点击要选择的列的任意单元格，然后点击确定。\n\n点击取消返回手动输入。",
                        "选择列",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.OK)
                    {
                        // 获取选择的单元格
                        try
                        {
                            object selection = app.GetType().InvokeMember("Selection", 
                                System.Reflection.BindingFlags.GetProperty, null, app, null);
                            
                            if (selection != null)
                            {
                                int column = (int)selection.GetType().InvokeMember("Column", 
                                    System.Reflection.BindingFlags.GetProperty, null, selection, null);
                                
                                if (column > 0)
                                {
                                    string columnLetter = ExcelHelper.GetColumnLetter(column);
                                    
                                    // 关键修复：正确设置文本框内容
                                    txtColumn.Text = columnLetter.ToUpper();
                                    
                                    // 显示窗体并聚焦
                                    this.Visible = true;
                                    this.TopMost = true;
                                    this.WindowState = FormWindowState.Normal;
                                    this.Activate();
                                    this.Focus();
                                    this.BringToFront();
                                    this.TopMost = false;
                                    
                                    // 确保文本框获得焦点并显示内容
                                    txtColumn.Focus();
                                    txtColumn.SelectAll();
                                    
                                    MessageBox.Show(string.Format("已选择列：{0}", columnLetter.ToUpper()), "选择成功", 
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    this.Visible = true;
                                    this.TopMost = true;
                                    this.Activate();
                                    this.Focus();
                                    this.TopMost = false;
                                    
                                    MessageBox.Show("未能获取选中的单元格，请重试或手动输入列号。", "选择失败", 
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    txtColumn.Focus();
                                }
                            }
                            else
                            {
                                this.Visible = true;
                                this.TopMost = true;
                                this.Activate();
                                this.Focus();
                                this.TopMost = false;
                                
                                MessageBox.Show("未能获取选中的单元格，请重试或手动输入列号。", "选择失败", 
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                txtColumn.Focus();
                            }
                        }
                        catch (Exception ex)
                        {
                            this.Visible = true;
                            this.TopMost = true;
                            this.Activate();
                            this.Focus();
                            this.TopMost = false;
                            
                            MessageBox.Show(string.Format("获取选中单元格时出错：{0}\n\n请手动输入列号。", ex.Message), 
                                "选择失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txtColumn.Focus();
                        }
                    }
                    else
                    {
                        // 用户取消
                        this.Visible = true;
                        this.TopMost = true;
                        this.Activate();
                        this.Focus();
                        this.TopMost = false;
                        txtColumn.Focus();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("操作WPS表格或Excel时出错：{0}", ex.Message), "错误", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("从表格选择列时发生错误：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 确保对话框总是可见
                if (!this.Visible)
                {
                    this.Visible = true;
                    this.TopMost = true;
                    this.Activate();
                    this.Focus();
                    this.TopMost = false;
                    txtColumn.Focus();
                }
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string input = txtColumn.Text.Trim().ToUpper();
            
            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show("请输入列号！", "输入验证", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtColumn.Focus();
                return;
            }

            if (!ExcelHelper.IsValidColumnLetter(input))
            {
                MessageBox.Show(string.Format("列号格式不正确：{0}\n\n请输入有效的列号，如A、B、AA、AB等。", input), 
                    "输入验证", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtColumn.Focus();
                txtColumn.SelectAll();
                return;
            }

            SelectedColumn = input;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
} 