// --- 文件 5: MatchForm.cs ---
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
        private bool isProcessing = false;
        private AppSettings settings;

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
                settings = AppSettings.Instance;
                ApplySettings();
                LoadMatcherSettings();
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
                Font newFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                this.AutoScaleMode = AutoScaleMode.Font;
                this.Font = newFont;
                this.menuStrip1.Font = newFont;
                this.PerformAutoScale();
            }
            catch (Exception ex)
            {
                WriteLog("应用通用设置失败: " + ex.Message, LogLevel.Warning);
            }
        }
        
        private void LoadMatcherSettings()
        {
            txtDelimiter.Text = settings.ConcatenationDelimiter;
            chkRemoveDuplicates.Checked = settings.RemoveDuplicateItems;
        }

        private void RefreshWorkbookList()
        {
            try
            {
                excelApp = ExcelAddin.GetExcelApplication();

                if (excelApp == null || !ExcelAddin.HasOpenWorkbooks(excelApp))
                {
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
            var displayNames = workbooks.Select(wb => wb.IsActive ? $"{wb.Name} [当前活动]" : wb.Name).ToArray();
            
            string prevShipping = cmbShippingWorkbook.SelectedItem?.ToString();
            string prevBill = cmbBillWorkbook.SelectedItem?.ToString();

            cmbShippingWorkbook.Items.Clear();
            cmbBillWorkbook.Items.Clear();
            cmbShippingWorkbook.Items.AddRange(displayNames);
            cmbBillWorkbook.Items.AddRange(displayNames);
            
            if (!string.IsNullOrEmpty(prevShipping) && cmbShippingWorkbook.Items.Contains(prevShipping))
                cmbShippingWorkbook.SelectedItem = prevShipping;
            else if (workbooks.Any(w => w.IsActive))
                cmbShippingWorkbook.SelectedIndex = workbooks.FindIndex(w => w.IsActive);
            else if (cmbShippingWorkbook.Items.Count > 0)
                cmbShippingWorkbook.SelectedIndex = 0;
            
            if (!string.IsNullOrEmpty(prevBill) && cmbBillWorkbook.Items.Contains(prevBill))
                cmbBillWorkbook.SelectedItem = prevBill;
            else if (workbooks.Any(w => w.IsActive))
                cmbBillWorkbook.SelectedIndex = workbooks.FindIndex(w => w.IsActive);
            else if (cmbBillWorkbook.Items.Count > 0)
                cmbBillWorkbook.SelectedIndex = 0;
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
        
        private void cmbShippingWorkbook_SelectedIndexChanged(object sender, EventArgs e) => LoadSheetsForWorkbook(cmbShippingWorkbook, cmbShippingSheet);
        private void cmbBillWorkbook_SelectedIndexChanged(object sender, EventArgs e) => LoadSheetsForWorkbook(cmbBillWorkbook, cmbBillSheet);
        
        private void cmbShippingSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateColumnComboBoxes(cmbShippingWorkbook, cmbShippingSheet, lblShippingInfo, cmbShippingTrackColumn, cmbShippingProductColumn, cmbShippingNameColumn);
        }
        
        private void cmbBillSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateColumnComboBoxes(cmbBillWorkbook, cmbBillSheet, lblBillInfo, cmbBillTrackColumn, cmbBillProductColumn, cmbBillNameColumn);
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

                string[] keywords = sheetCombo == cmbShippingSheet ? new[] { "发货明细", "发货" } : new[] { "账单明细", "账单" };
                SetDefaultSheet(sheetCombo, keywords);
            }
            catch (Exception ex)
            {
                WriteLog("加载工作表失败: " + ex.Message, LogLevel.Error);
            }
        }
        
        private void PopulateColumnComboBoxes(ComboBox wbCombo, ComboBox wsCombo, Label infoLabel, params ComboBox[] columnCombos)
        {
            foreach (var combo in columnCombos) { combo.DataSource = null; combo.Items.Clear(); combo.Text = ""; }
            infoLabel.Text = "";
            if (wbCombo.SelectedIndex < 0 || wsCombo.SelectedIndex < 0 || wsCombo.SelectedItem == null) return;

            try
            {
                var wbInfo = workbooks[wbCombo.SelectedIndex];
                var ws = wbInfo.Workbook.Worksheets[wsCombo.SelectedItem.ToString()] as Excel.Worksheet;
                if (ws == null) return;

                var usedRange = ws.UsedRange;
                if (usedRange.Rows.Count == 0) return;

                var fileInfo = new FileInfo(wbInfo.Workbook.FullName);
                infoLabel.Text = $"总行数: {usedRange.Rows.Count:N0} | 文件大小: {(double)fileInfo.Length / (1024 * 1024):F2} MB";

                int colCount = usedRange.Columns.Count;
                var headerRow = FindHeaderRow(usedRange);
                var headers = headerRow?.Value2 as object[,];

                var columnItems = new List<ColumnInfo>();
                if (headers != null)
                {
                    for (int i = 1; i <= colCount; i++)
                    {
                        string colLetter = ExcelHelper.GetColumnLetter(i);
                        string headerText = headers[1, i]?.ToString().Trim() ?? "";
                        string previewData = GetColumnPreviewData(ws, i, usedRange.Rows.Count);
                        
                        if (headerText.Length > 15) headerText = headerText.Substring(0, 15) + "...";
                        if (previewData.Length > 20) previewData = previewData.Substring(0, 20) + "...";
                        
                        columnItems.Add(new ColumnInfo
                        {
                            DisplayText = $"{colLetter} ({headerText})",
                            ColumnLetter = colLetter,
                            HeaderText = headerText,
                            PreviewData = previewData,
                            SearchKeywords = $"{colLetter} {headerText} {previewData}"
                        });
                    }
                }
                
                foreach (var combo in columnCombos)
                {
                    combo.DisplayMember = "DisplayText";
                    combo.ValueMember = "ColumnLetter";
                    combo.DataSource = new BindingSource(columnItems, null);
                    combo.SelectedIndex = -1;
                    
                    // 启用自动完成和搜索功能
                    combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    combo.AutoCompleteSource = AutoCompleteSource.ListItems;
                    
                    // 添加搜索功能
                    combo.TextChanged += (s, e) => FilterComboBoxItems(combo, columnItems);
                }
                
                // 智能选择默认列
                AutoSelectDefaultColumns(columnCombos, columnItems);
            }
            catch (Exception ex)
            {
                 WriteLog("填充列下拉框失败: " + ex.Message, LogLevel.Error);
            }
        }
        
        private string GetColumnPreviewData(Excel.Worksheet ws, int columnIndex, int totalRows)
        {
            try
            {
                // 从第二行开始查找非空数据作为预览
                for (int row = 2; row <= Math.Min(totalRows, 100); row++)
                {
                    var cell = ws.Cells[row, columnIndex] as Excel.Range;
                    if (cell != null && cell.Value2 != null)
                    {
                        string value = cell.Value2.ToString().Trim();
                        if (!string.IsNullOrEmpty(value))
                        {
                            return value;
                        }
                    }
                }
                return "无数据";
            }
            catch
            {
                return "无数据";
            }
        }
        
        private void FilterComboBoxItems(ComboBox combo, List<ColumnInfo> allItems)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(combo.Text))
                {
                    combo.DataSource = new BindingSource(allItems, null);
                    return;
                }
                
                string searchText = combo.Text.ToLower();
                var filteredItems = allItems.Where(item => 
                    item.SearchKeywords.ToLower().Contains(searchText) ||
                    item.ColumnLetter.ToLower().Contains(searchText) ||
                    item.HeaderText.ToLower().Contains(searchText) ||
                    item.PreviewData.ToLower().Contains(searchText)
                ).ToList();
                
                combo.DataSource = new BindingSource(filteredItems, null);
            }
            catch (Exception ex)
            {
                WriteLog("过滤列项目失败: " + ex.Message, LogLevel.Warning);
            }
        }
        
        private void AutoSelectDefaultColumns(ComboBox[] columnCombos, List<ColumnInfo> columnItems)
        {
            try
            {
                // 运单号列智能选择
                var trackColumn = columnItems.FirstOrDefault(item => 
                    item.HeaderText.Contains("运单") || 
                    item.HeaderText.Contains("快递") || 
                    item.HeaderText.Contains("单号") ||
                    item.HeaderText.Contains("track") ||
                    item.HeaderText.Contains("tracking"));
                    
                if (trackColumn != null && columnCombos.Length > 0)
                {
                    columnCombos[0].SelectedValue = trackColumn.ColumnLetter;
                }
                
                // 商品编码列智能选择
                if (columnCombos.Length > 1)
                {
                    var productCodeColumn = columnItems.FirstOrDefault(item => 
                        item.HeaderText.Contains("商品") && 
                        (item.HeaderText.Contains("编码") || item.HeaderText.Contains("代码") || item.HeaderText.Contains("code")));
                        
                    if (productCodeColumn != null)
                    {
                        columnCombos[1].SelectedValue = productCodeColumn.ColumnLetter;
                    }
                }
                
                // 商品名称列智能选择
                if (columnCombos.Length > 2)
                {
                    var productNameColumn = columnItems.FirstOrDefault(item => 
                        item.HeaderText.Contains("商品") && 
                        (item.HeaderText.Contains("名称") || item.HeaderText.Contains("品名") || item.HeaderText.Contains("name")));
                        
                    if (productNameColumn != null)
                    {
                        columnCombos[2].SelectedValue = productNameColumn.ColumnLetter;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("自动选择默认列失败: " + ex.Message, LogLevel.Warning);
            }
        }
        
        private Excel.Range FindHeaderRow(Excel.Range usedRange)
        {
            for (int i = 1; i <= Math.Min(100, usedRange.Rows.Count); i++)
            {
                var row = usedRange.Rows[i] as Excel.Range;
                var rowData = row.Value2 as object[,];
                if (rowData != null && Enumerable.Range(1, rowData.GetLength(1)).Any(col => rowData[1, col] != null))
                {
                    return row;
                }
            }
            return usedRange.Rows[1] as Excel.Range;
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
                settings.ConcatenationDelimiter = txtDelimiter.Text;
                settings.RemoveDuplicateItems = chkRemoveDuplicates.Checked;
                settings.Save();
                
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
            this.isProcessing = processing;
            menuStrip1.Enabled = !processing;
            tabControlMain.Enabled = !processing;
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
                RefreshWorkbookList();
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

                if (!AreColumnsValid(shippingWb.Workbook, cmbShippingSheet.Text, "发货", cmbShippingTrackColumn, cmbShippingProductColumn, cmbShippingNameColumn)) return false;
                if (!AreColumnsValid(billWb.Workbook, cmbBillSheet.Text, "账单", cmbBillTrackColumn, cmbBillProductColumn, cmbBillNameColumn)) return false;
                
                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show($"验证列时发生错误: {ex.Message}", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }
        


        private bool AreColumnsValid(Excel.Workbook wb, string sheetName, string type, params ComboBox[] columnCombos)
        {
            if (wb.Worksheets[sheetName] is Excel.Worksheet ws)
            {
                foreach (var cb in columnCombos)
                {
                    if (cb.SelectedItem == null)
                    {
                        MessageBox.Show($"请为\"{type}\"表选择列！", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cb.Focus();
                        return false;
                    }
                    
                    var selectedItem = cb.SelectedItem as ColumnInfo;
                    if (selectedItem == null)
                    {
                        MessageBox.Show($"您为\"{type}\"表选择的列无效。", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cb.Focus();
                        return false;
                    }
                    
                    // 验证列是否存在于工作表中
                    if (!ExcelHelper.IsValidColumnLetter(selectedItem.ColumnLetter))
                    {
                        MessageBox.Show($"您为\"{type}\"表选择的列\"{selectedItem.DisplayText}\"无效。", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cb.Focus();
                        return false;
                    }
                }
                return true;
            }
            return false;
        }

        private MultiWorkbookMatchConfig CreateMatchConfig()
        {
            return new MultiWorkbookMatchConfig
            {
                ShippingWorkbook = workbooks[cmbShippingWorkbook.SelectedIndex].Workbook,
                BillWorkbook = workbooks[cmbBillWorkbook.SelectedIndex].Workbook,
                ShippingSheetName = cmbShippingSheet.SelectedItem.ToString(),
                BillSheetName = cmbBillSheet.SelectedItem.ToString(),
                ShippingTrackColumn = (cmbShippingTrackColumn.SelectedItem as ColumnInfo)?.ColumnLetter ?? "",
                ShippingProductColumn = (cmbShippingProductColumn.SelectedItem as ColumnInfo)?.ColumnLetter ?? "",
                ShippingNameColumn = (cmbShippingNameColumn.SelectedItem as ColumnInfo)?.ColumnLetter ?? "",
                BillTrackColumn = (cmbBillTrackColumn.SelectedItem as ColumnInfo)?.ColumnLetter ?? "",
                BillProductColumn = (cmbBillProductColumn.SelectedItem as ColumnInfo)?.ColumnLetter ?? "",
                BillNameColumn = (cmbBillNameColumn.SelectedItem as ColumnInfo)?.ColumnLetter ?? ""
            };
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var config = e.Argument as MultiWorkbookMatchConfig;
            var service = new MatchService();
            service.CancellationCheck = () => backgroundWorker.CancellationPending;
            
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
                    lblStatus.Text = result.ErrorMessage == "任务被用户取消" ? "任务已由用户停止。" : "匹配失败！";
                    if(result.ErrorMessage != "任务被用户取消")
                    {
                        MessageBox.Show($"匹配失败：{result.ErrorMessage}\n\n请查看日志获取详细信息。", "匹配失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
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
            string aboutInfo = "YY 运单匹配工具 v3.0 (智能版)\n\n" +
                             "功能特点：\n" +
                             "• 智能运单匹配，支持灵活拼接\n" +
                             "• 支持多工作簿操作与动态加载\n" +
                             "• 高级列选择(带预览和智能搜索)\n" +
                             "• 智能默认列选择\n" +
                             "• 优化的用户界面和性能\n\n" +
                             "作者: 皮皮熊\n" +
                             "邮箱: oyxo@qq.com";
            MessageBox.Show(aboutInfo, "关于 YY工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (isProcessing)
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
            if (isProcessing)
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