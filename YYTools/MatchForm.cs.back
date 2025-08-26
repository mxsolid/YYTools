using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    public partial class MatchForm : Form
    {
        private Excel.Application excelApp;
        private BackgroundWorker backgroundWorker;
        private bool isProcessing = false;
        private List<Excel.Workbook> workbooks;
        
        public MatchForm()
        {
            InitializeComponent();
            InitializeBackgroundWorker();
            LoadWorkbooks();
            ApplySettings();
        }

        private void InitializeBackgroundWorker()
        {
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }

        private void LoadWorkbooks()
        {
            try
            {
                excelApp = ExcelAddin.GetExcelApplication();
                if (excelApp == null)
                {
                    MessageBox.Show("请先打开WPS表格或Excel文件！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                workbooks = ExcelAddin.GetWorkbooks();
                
                cmbBillWorkbook.Items.Clear();
                cmbShippingWorkbook.Items.Clear();
                
                foreach (var workbook in workbooks)
                {
                    cmbBillWorkbook.Items.Add(workbook.Name);
                    cmbShippingWorkbook.Items.Add(workbook.Name);
                }

                if (cmbBillWorkbook.Items.Count > 0)
                {
                    cmbBillWorkbook.SelectedIndex = 0;
                    cmbShippingWorkbook.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("加载工作簿失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplySettings()
        {
            try
            {
                var settings = AppSettings.Instance;
                
                // 应用字体设置
                Font newFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                ApplyFontToAllControls(this, newFont);
                
                // 应用界面缩放
                if (settings.AutoScaleUI)
                {
                    this.AutoScaleMode = AutoScaleMode.Dpi;
                }
            }
            catch
            {
                // 设置应用失败时使用默认值
            }
        }

        private void ApplyFontToAllControls(Control parent, Font font)
        {
            foreach (Control control in parent.Controls)
            {
                control.Font = font;
                if (control.HasChildren)
                {
                    ApplyFontToAllControls(control, font);
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (isProcessing)
            {
                MessageBox.Show("任务正在进行中，请等待完成！", "提示", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 验证选择
                if (cmbBillWorkbook.SelectedIndex < 0 || cmbShippingWorkbook.SelectedIndex < 0)
                {
                    MessageBox.Show("请选择工作簿！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (cmbBillSheet.SelectedIndex < 0 || cmbShippingSheet.SelectedIndex < 0)
                {
                    MessageBox.Show("请选择工作表！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 准备配置
                var config = new MatchConfig
                {
                    BillWorkbook = workbooks[cmbBillWorkbook.SelectedIndex],
                    ShippingWorkbook = workbooks[cmbShippingWorkbook.SelectedIndex],
                    BillSheetName = cmbBillSheet.SelectedItem.ToString(),
                    ShippingSheetName = cmbShippingSheet.SelectedItem.ToString(),
                    BillTrackColumn = int.Parse(txtBillTrackColumn.Text),
                    BillProductColumn = int.Parse(txtBillProductColumn.Text),
                    BillNameColumn = int.Parse(txtBillNameColumn.Text),
                    ShippingTrackColumn = int.Parse(txtShippingTrackColumn.Text),
                    ShippingProductColumn = int.Parse(txtShippingProductColumn.Text),
                    ShippingNameColumn = int.Parse(txtShippingNameColumn.Text)
                };

                // 开始处理
                isProcessing = true;
                btnStart.Enabled = false;
                btnStart.Text = "处理中...";
                progressBar.Visible = true;
                progressBar.Value = 0;

                backgroundWorker.RunWorkerAsync(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动处理失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                isProcessing = false;
                btnStart.Enabled = true;
                btnStart.Text = "开始匹配";
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var config = (MatchConfig)e.Argument;
                var service = new MatchService();
                
                var result = service.ExecuteMatch(config, (progress, message) =>
                {
                    backgroundWorker.ReportProgress(progress, message);
                });
                
                e.Result = result;
            }
            catch (Exception ex)
            {
                e.Result = new MatchResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            lblStatus.Text = e.UserState != null ? e.UserState.ToString() : "";
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            isProcessing = false;
            btnStart.Enabled = true;
            btnStart.Text = "开始匹配";
            progressBar.Visible = false;

            if (e.Result is MatchResult result)
            {
                ShowResult(result);
            }
        }

        private void ShowResult(MatchResult result)
        {
            if (result.Success)
            {
                string message = string.Format(
                    "🎉 匹配完成！\n\n" +
                    "📊 处理统计：\n" +
                    "• 处理行数：{0:N0}\n" +
                    "• 匹配数量：{1:N0}\n" +
                    "• 填充单元格：{2:N0}\n" +
                    "• 处理时间：{3:F2} 秒\n" +
                    "• 处理速度：{4:F0} 行/秒\n\n" +
                    "✅ 数据已成功写入账单明细表！",
                    result.ProcessedRows,
                    result.MatchedCount,
                    result.UpdatedCells,
                    result.ElapsedSeconds,
                    result.ProcessedRows / Math.Max(result.ElapsedSeconds, 0.001)
                );
                
                MessageBox.Show(message, "成功", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                string message = string.Format(
                    "❌ 匹配失败\n\n" +
                    "错误信息：{0}\n\n" +
                    "请检查：\n" +
                    "• 工作表和列设置是否正确\n" +
                    "• 数据格式是否符合要求\n" +
                    "• 文件是否可以正常访问",
                    result.ErrorMessage
                );
                
                MessageBox.Show(message, "失败", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (isProcessing)
            {
                DialogResult result = MessageBox.Show(
                    "确定要停止当前任务吗？", "确认", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    backgroundWorker.CancelAsync();
                    this.Close();
                }
            }
            else
            {
                this.Close();
            }
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            try
            {
                var settingsForm = new SettingsForm();
                if (settingsForm.ShowDialog() == DialogResult.OK)
                {
                    ApplySettings();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开设置失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbBillWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbBillWorkbook, cmbBillSheet);
        }

        private void cmbShippingWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbShippingWorkbook, cmbShippingSheet);
        }

        private void LoadSheetsForWorkbook(ComboBox workbookCombo, ComboBox sheetCombo)
        {
            try
            {
                if (workbooks == null || workbookCombo.SelectedIndex < 0) return;

                var selectedWorkbook = workbooks[workbookCombo.SelectedIndex];
                var sheetNames = ExcelAddin.GetWorksheetNames(selectedWorkbook);
                
                sheetCombo.Items.Clear();
                foreach (string sheetName in sheetNames)
                {
                    sheetCombo.Items.Add(sheetName);
                }

                if (sheetCombo.Items.Count > 0)
                {
                    sheetCombo.SelectedIndex = 0;
                }

                // 强制刷新界面
                sheetCombo.Refresh();
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show("加载工作表失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    public class MatchConfig
    {
        public Excel.Workbook BillWorkbook { get; set; }
        public Excel.Workbook ShippingWorkbook { get; set; }
        public string BillSheetName { get; set; }
        public string ShippingSheetName { get; set; }
        public int BillTrackColumn { get; set; }
        public int BillProductColumn { get; set; }
        public int BillNameColumn { get; set; }
        public int ShippingTrackColumn { get; set; }
        public int ShippingProductColumn { get; set; }
        public int ShippingNameColumn { get; set; }
    }

    public class MatchResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public int ProcessedRows { get; set; }
        public int MatchedCount { get; set; }
        public int UpdatedCells { get; set; }
        public double ElapsedSeconds { get; set; }
    }
}
