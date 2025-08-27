using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    public partial class MatchForm : Form
    {
        private Excel.Application excelApp;
        private BackgroundWorker backgroundWorker;
        private List<WorkbookInfo> workbooks = new List<WorkbookInfo>();
        private bool isProcessing = false; // <<< --- 修复: 重新声明缺失的字段

        public MatchForm()
        {
            InitializeComponent();
            InitializeCustomComponents();
            InitializeBackgroundWorker();
            InitializeForm();
        }

        private void InitializeCustomComponents()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;
            this.Shown += (s, e) => { this.Activate(); };
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

        private void InitializeForm()
        {
            try
            {
                ApplySettings();
                SetDefaultValues();
                RefreshWorkbookList();
                MatchService.CleanupOldLogs();
            }
            catch (Exception ex)
            {
                WriteLog("初始化窗体失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show("初始化失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplySettings()
        {
            try
            {
                AppSettings settings = AppSettings.Instance;
                Font newFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                this.AutoScaleMode = AutoScaleMode.Font;
                this.Font = newFont;
                this.PerformAutoScale();
            }
            catch (Exception ex)
            {
                WriteLog("应用设置失败: " + ex.Message, LogLevel.Warning);
            }
        }

        private void RefreshWorkbookList()
        {
            try
            {
                WriteLog("开始加载工作簿列表", LogLevel.Info);
                excelApp = ExcelAddin.GetExcelApplication();

                if (excelApp == null || !ExcelAddin.HasOpenWorkbooks(excelApp))
                {
                    WriteLog("没有检测到Excel/WPS进程或打开的工作簿", LogLevel.Warning);
                    UpdateUIForNoWorkbooks();
                    return;
                }

                workbooks = ExcelAddin.GetOpenWorkbooks();

                if (workbooks.Count == 0)
                {
                    UpdateUIForNoWorkbooks();
                    return;
                }

                UpdateUIWithWorkbooks();
                PopulateComboBoxes();
            }
            catch (Exception ex)
            {
                WriteLog("加载工作簿失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show("加载工作簿失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateUIForNoWorkbooks();
            }
        }

        private void PopulateComboBoxes()
        {
            string prevShippingWb = cmbShippingWorkbook.Text;
            string prevBillWb = cmbBillWorkbook.Text;

            cmbShippingWorkbook.Items.Clear();
            cmbBillWorkbook.Items.Clear();

            var displayNames = workbooks.Select(wb => wb.IsActive ? $"{wb.Name} [当前活动]" : wb.Name).ToArray();
            cmbShippingWorkbook.Items.AddRange(displayNames);
            cmbBillWorkbook.Items.AddRange(displayNames);
            
            int activeIndex = workbooks.FindIndex(wb => wb.IsActive);
            if (activeIndex != -1)
            {
                cmbShippingWorkbook.SelectedIndex = activeIndex;
                cmbBillWorkbook.SelectedIndex = activeIndex;
            }
            else if (workbooks.Count > 0)
            {
                cmbShippingWorkbook.SelectedIndex = 0;
                cmbBillWorkbook.SelectedIndex = 0;
            }
        }

        private void UpdateUIForNoWorkbooks()
        {
            gbShipping.Enabled = false;
            gbBill.Enabled = false;
            btnStart.Enabled = false;
            lblStatus.Text = "未检测到打开的Excel/WPS文件。请打开文件或从菜单栏选择文件。";
        }

        private void UpdateUIWithWorkbooks()
        {
            gbShipping.Enabled = true;
            gbBill.Enabled = true;
            btnStart.Enabled = true;
            lblStatus.Text = $"已加载 {workbooks.Count} 个工作簿。请配置并开始任务。";
        }
        
        private void cmbShippingWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbShippingWorkbook, cmbShippingSheet);
            UpdateWorkbookInfo(cmbShippingWorkbook, lblShippingInfo);
        }

        private void cmbBillWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbBillWorkbook, cmbBillSheet);
            UpdateWorkbookInfo(cmbBillWorkbook, lblBillInfo);
        }

        private void LoadSheetsForWorkbook(ComboBox workbookCombo, ComboBox sheetCombo)
        {
            sheetCombo.Items.Clear();
            if (workbookCombo.SelectedIndex < 0 || workbookCombo.SelectedIndex >= workbooks.Count) return;

            try
            {
                Excel.Workbook selectedWorkbook = workbooks[workbookCombo.SelectedIndex].Workbook;
                List<string> sheetNames = ExcelAddin.GetWorksheetNames(selectedWorkbook);
                sheetCombo.Items.AddRange(sheetNames.ToArray());

                string[] keywords = sheetCombo == cmbShippingSheet
                    ? new[] { "发货明细", "发货" }
                    : new[] { "账单明细", "账单" };
                SetDefaultSheet(sheetCombo, keywords);
            }
            catch (Exception ex)
            {
                WriteLog("加载工作表失败: " + ex.Message, LogLevel.Error);
            }
        }

        private void UpdateWorkbookInfo(ComboBox workbookCombo, Label infoLabel)
        {
            infoLabel.Text = "总行数: - | 文件大小: -";
            if (workbookCombo.SelectedIndex < 0 || workbookCombo.SelectedIndex >= workbooks.Count) return;

            try
            {
                WorkbookInfo wbInfo = workbooks[workbookCombo.SelectedIndex];
                Excel.Worksheet activeSheet = (Excel.Worksheet)wbInfo.Workbook.ActiveSheet;
                int rowCount = activeSheet.UsedRange.Rows.Count;

                var fileInfo = new FileInfo(wbInfo.Workbook.FullName);
                double fileSizeMB = (double)fileInfo.Length / (1024 * 1024);

                infoLabel.Text = $"总行数: {rowCount:N0} | 文件大小: {fileSizeMB:F2} MB";
            }
            catch (Exception ex)
            {
                 WriteLog("更新工作簿信息失败: " + ex.Message, LogLevel.Warning);
            }
        }

        private void SetDefaultValues()
        {
            try
            {
                AppSettings settings = AppSettings.Instance;
                txtShippingTrackColumn.Text = settings.DefaultShippingTrackColumn;
                txtShippingProductColumn.Text = settings.DefaultShippingProductColumn;
                txtShippingNameColumn.Text = settings.DefaultShippingNameColumn;
                txtBillTrackColumn.Text = settings.DefaultBillTrackColumn;
                txtBillProductColumn.Text = settings.DefaultBillProductColumn;
                txtBillNameColumn.Text = settings.DefaultBillNameColumn;
            }
            catch (Exception ex)
            {
                WriteLog("设置默认值失败: " + ex.Message, LogLevel.Warning);
            }
        }

        private void SetDefaultSheet(ComboBox combo, string[] keywords)
        {
            if (combo.Items.Count == 0) return;
            foreach (string item in combo.Items)
            {
                if (keywords.Any(k => item.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    combo.SelectedItem = item;
                    return;
                }
            }
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!ValidateInput() || !ValidateColumns()) return;

            try
            {
                MultiWorkbookMatchConfig config = CreateMatchConfig();
                SetUiProcessingState(true);
                backgroundWorker.RunWorkerAsync(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动匹配失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetUiProcessingState(false);
            }
        }

        private void SetUiProcessingState(bool processing)
        {
            this.isProcessing = processing; // 修复: 使用 this.isProcessing 访问字段
            menuStrip1.Enabled = !processing;
            gbShipping.Enabled = !processing;
            gbBill.Enabled = !processing;
            btnStart.Enabled = !processing;
            progressBar.Visible = processing;

            if (processing)
            {
                progressBar.Value = 0;
                lblStatus.Text = "正在初始化匹配任务...";
                btnClose.Text = "⏹️ 停止任务";
            }
            else
            {
                btnClose.Text = "关闭";
            }
        }

        private bool ValidateInput()
        {
            if (cmbShippingWorkbook.SelectedIndex < 0 || cmbBillWorkbook.SelectedIndex < 0)
            {
                MessageBox.Show("请选择工作簿！", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (cmbShippingSheet.SelectedIndex < 0 || cmbBillSheet.SelectedIndex < 0)
            {
                MessageBox.Show("请选择工作表！", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private bool ValidateColumns()
        {
            try
            {
                var shippingWb = workbooks[cmbShippingWorkbook.SelectedIndex];
                var billWb = workbooks[cmbBillWorkbook.SelectedIndex];

                if (!AreColumnsValid(shippingWb.Workbook, cmbShippingSheet.Text, "发货", txtShippingTrackColumn, txtShippingProductColumn, txtShippingNameColumn)) return false;
                if (!AreColumnsValid(billWb.Workbook, cmbBillSheet.Text, "账单", txtBillTrackColumn, txtBillProductColumn, txtBillNameColumn)) return false;
                
                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show($"验证列时发生错误: {ex.Message}", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private bool AreColumnsValid(Excel.Workbook wb, string sheetName, string type, params TextBox[] columnTextBoxes)
        {
            Excel.Worksheet ws = wb.Worksheets[sheetName] as Excel.Worksheet;
            if (ws == null) return false;
            int maxCols = ws.UsedRange.Columns.Count;

            foreach (var tb in columnTextBoxes)
            {
                if (!ExcelHelper.IsValidColumnLetter(tb.Text))
                {
                     MessageBox.Show($"您为“{type}”表输入的列名“{tb.Text}”格式无效。", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                     return false;
                }
                int colIndex = ExcelHelper.GetColumnNumber(tb.Text);
                if (colIndex > maxCols)
                {
                    MessageBox.Show($"您为“{type}”表指定的列“{tb.Text}”超出了工作表的最大列范围({ExcelHelper.GetColumnLetter(maxCols)})。", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            return true;
        }

        private MultiWorkbookMatchConfig CreateMatchConfig()
        {
            return new MultiWorkbookMatchConfig
            {
                ShippingWorkbook = workbooks[cmbShippingWorkbook.SelectedIndex].Workbook,
                BillWorkbook = workbooks[cmbBillWorkbook.SelectedIndex].Workbook,
                ShippingSheetName = cmbShippingSheet.SelectedItem.ToString(),
                BillSheetName = cmbBillSheet.SelectedItem.ToString(),
                ShippingTrackColumn = txtShippingTrackColumn.Text,
                ShippingProductColumn = txtShippingProductColumn.Text,
                ShippingNameColumn = txtShippingNameColumn.Text,
                BillTrackColumn = txtBillTrackColumn.Text,
                BillProductColumn = txtBillProductColumn.Text,
                BillNameColumn = txtBillNameColumn.Text
            };
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var config = e.Argument as MultiWorkbookMatchConfig;
            var service = new MatchService();
            e.Result = service.ExecuteMatch(config, (progress, message) => {
                if (backgroundWorker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                backgroundWorker.ReportProgress(progress, message);
            });
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = Math.Min(e.ProgressPercentage, 100);
            lblStatus.Text = e.UserState?.ToString();
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SetUiProcessingState(false);
            if (e.Cancelled)
            {
                lblStatus.Text = "任务已由用户停止。";
                return;
            }
            if (e.Error != null)
            {
                lblStatus.Text = "处理出错！";
                MessageBox.Show($"处理过程中发生错误：{e.Error.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (e.Result is MatchResult result)
            {
                if (!result.Success)
                {
                    lblStatus.Text = "匹配失败！";
                    MessageBox.Show($"匹配失败：{result.ErrorMessage}\n\n请查看日志获取详细信息。", "匹配失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    lblStatus.Text = $"🎉 任务完成！耗时 {result.ElapsedSeconds:F2} 秒";
                    ShowResultDialog(result);
                }
            }
        }
        
        private void ShowResultDialog(MatchResult result)
        {
             string summary = result.MatchedCount == 0
                ? $"匹配完成，但没有找到匹配的运单！\n\n处理的账单行数：{result.ProcessedRows:N0}\n处理耗时：{result.ElapsedSeconds:F2} 秒"
                : $"🎉 运单匹配任务完成！\n================================\n\n" +
                  $"📊 处理统计：\n" +
                  $"  • 处理账单行数：{result.ProcessedRows:N0} 行\n" +
                  $"  • 成功匹配运单：{result.MatchedCount:N0} 个\n" +
                  $"  • 填充数据单元格：{result.UpdatedCells:N0} 个\n\n" +
                  $"⚡ 性能表现：\n" +
                  $"  • 总处理时间：{result.ElapsedSeconds:F2} 秒";

            MessageBox.Show(summary, "任务完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // --- Menu Click Handlers ---
        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel 工作簿 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
                ofd.Title = "请选择一个Excel文件";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var openedWb = ExcelAddin.LoadWorkbookFromFile(ofd.FileName);
                        if (openedWb != null) RefreshWorkbookList();
                        else MessageBox.Show("无法打开指定的文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"打开文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void refreshListToolStripMenuItem_Click(object sender, EventArgs e) => RefreshWorkbookList();
        private void exitToolStripMenuItem_Click(object sender, EventArgs e) => this.Close();
        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
             try
            {
                using (SettingsForm settingsForm = new SettingsForm())
                {
                    if (settingsForm.ShowDialog() == DialogResult.OK)
                    {
                        ApplySettings();
                        SetDefaultValues();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开设置窗口失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void viewLogsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string logPath = MatchService.GetLogFolderPath();
                if (Directory.Exists(logPath))
                {
                    System.Diagnostics.Process.Start("explorer.exe", logPath);
                }
                else
                {
                    MessageBox.Show("日志文件夹不存在。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开日志文件夹时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string aboutInfo = "YY 运单匹配工具 v2.3\n\n" +
                             "功能特点：\n" +
                             "• 智能运单匹配，支持多性能模式\n" +
                             "• 支持多工作簿操作与动态加载\n" +
                             "• 自动列选择和有效性验证\n\n" +
                             "适用于：WPS表格、Microsoft Excel";
            MessageBox.Show(aboutInfo, "关于 YY工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        // --- Column Selection Button Handlers ---
        private void btnSelectTrackCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingTrackColumn, "发货明细运单号列");
        private void btnSelectProductCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingProductColumn, "发货明细商品编码列");
        private void btnSelectNameCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingNameColumn, "发货明细商品名称列");
        private void btnSelectBillTrackCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbBillWorkbook, txtBillTrackColumn, "账单明细运单号列");
        private void btnSelectBillProductCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbBillWorkbook, txtBillProductColumn, "账单明细商品编码列");
        private void btnSelectBillNameCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbBillWorkbook, txtBillNameColumn, "账单明细商品名称列");
        
        private void SelectColumnForWorkbook(ComboBox workbookCombo, TextBox targetTextBox, string title)
        {
            if (workbookCombo.SelectedIndex < 0) return;
            try
            {
                this.Visible = false;
                Excel.Workbook selectedWorkbook = workbooks[workbookCombo.SelectedIndex].Workbook;
                selectedWorkbook.Activate();
                
                MessageBox.Show($"请在工作簿 [{selectedWorkbook.Name}] 中选择 {title} 所在的任意一个单元格，然后点击“确定”。", "选择列", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (selectedWorkbook.Application.Selection is Excel.Range selection)
                {
                    targetTextBox.Text = ExcelHelper.GetColumnLetter(selection.Column);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择列时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Visible = true;
                this.Activate();
            }
        }
        
        private void btnClose_Click(object sender, EventArgs e)
        {
            if (isProcessing) // 修复: 使用 isProcessing 字段
            {
                if (backgroundWorker.IsBusy)
                {
                    backgroundWorker.CancelAsync();
                }
            }
            else
            {
                this.Close();
            }
        }

        private void WriteLog(string message, LogLevel level) => MatchService.WriteLog($"[MatchForm] {message}", level);

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (isProcessing) // 修复: 使用 isProcessing 字段
            {
                if (MessageBox.Show("任务正在处理中，确定要强制退出吗？", "确认退出", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    if (backgroundWorker.IsBusy) backgroundWorker.CancelAsync();
                }
            }
            base.OnFormClosing(e);
        }
    }
}