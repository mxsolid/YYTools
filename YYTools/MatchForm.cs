using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// 运单匹配配置窗体 - 支持多工作簿
    /// </summary>
    public partial class MatchForm : Form
    {
        private Excel.Application excelApp;
        private BackgroundWorker backgroundWorker;
        private bool isProcessing = false;
        private List<WorkbookInfo> workbooks;

        public MatchForm()
        {
            InitializeComponent();
            InitializeCustomComponents();
            InitializeBackgroundWorker();
            
            InitializeForm();
        }

        /// <summary>
        /// 初始化自定义组件和窗体属性
        /// </summary>
        private void InitializeCustomComponents()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;

            this.Shown += (s, e) => {
                this.Activate(); // 确保窗体显示时获得焦点
            };
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
                
                excelApp = ExcelAddin.GetExcelApplication();
                
                LoadWorkbooks(); // LoadWorkbooks现在会处理没有Excel实例的情况
                
                SetDefaultValues();
                MatchService.CleanupOldLogs();
            }
            catch (Exception ex)
            {
                WriteLog("初始化窗体失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show("初始化失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 应用设置到窗体，并强制重新缩放以适应字体
        /// </summary>
        private void ApplySettings()
        {
            try
            {
                AppSettings settings = AppSettings.Instance;
                Font newFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                
                // 将AutoScaleMode设置为Font是实现字体动态缩放的关键
                this.AutoScaleMode = AutoScaleMode.Font;
                this.Font = newFont; // 应用基础字体
                ApplyFontToControls(this, newFont);

                // 强制窗体根据新字体重新计算布局，解决控件内容被截断的问题
                this.PerformAutoScale(); 
            }
            catch (Exception ex)
            {
                WriteLog("应用设置失败: " + ex.Message, LogLevel.Warning);
            }
        }
        
        private void ApplyFontToControls(Control parent, Font font)
        {
            foreach (Control control in parent.Controls)
            {
                control.Font = font;
                if (control.HasChildren)
                {
                    ApplyFontToControls(control, font);
                }
            }
        }
        
        /// <summary>
        /// 加载工作簿列表 - 已修改为不退出程序
        /// </summary>
        private void LoadWorkbooks()
        {
            try
            {
                WriteLog("开始加载工作簿列表", LogLevel.Info);
                
                excelApp = ExcelAddin.GetExcelApplication();
                if (excelApp == null || !ExcelAddin.HasOpenWorkbooks(excelApp))
                {
                    WriteLog("没有检测到Excel/WPS进程或打开的工作簿", LogLevel.Warning);
                    SetUIForNoWorkbooksState(); // 进入无工作簿状态
                    return;
                }

                workbooks = ExcelAddin.GetOpenWorkbooks();
                
                if (workbooks == null || workbooks.Count == 0)
                {
                    WriteLog("再次确认没有检测到打开的工作簿", LogLevel.Warning);
                    SetUIForNoWorkbooksState(); // 确认无工作簿，进入相应状态
                    return;
                }
                
                // 如果成功加载，恢复UI
                RestoreUIState();
                
                cmbShippingWorkbook.Items.Clear();
                cmbBillWorkbook.Items.Clear();
                
                WriteLog("检测到 " + workbooks.Count + " 个工作簿", LogLevel.Info);
                
                int activeIndex = workbooks.FindIndex(wb => wb.IsActive);
                
                for (int i = 0; i < workbooks.Count; i++)
                {
                    var workbookInfo = workbooks[i];
                    string displayName = workbookInfo.Name;
                    
                    if (workbookInfo.IsActive)
                    {
                        displayName += " [当前活动]";
                    }
                    
                    cmbShippingWorkbook.Items.Add(displayName);
                    cmbBillWorkbook.Items.Add(displayName);
                }
                
                if (activeIndex >= 0)
                {
                    cmbShippingWorkbook.SelectedIndex = activeIndex;
                    cmbBillWorkbook.SelectedIndex = activeIndex;
                }
                else if (workbooks.Count > 0)
                {
                    cmbShippingWorkbook.SelectedIndex = 0;
                    cmbBillWorkbook.SelectedIndex = 0;
                }
                
                lblStatus.Text = $"已加载 {workbooks.Count} 个工作簿。";
                WriteLog("工作簿加载完成", LogLevel.Info);
            }
            catch (Exception ex)
            {
                WriteLog("加载工作簿失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show("加载工作簿失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetUIForNoWorkbooksState();
            }
        }

        /// <summary>
        /// 当没有工作簿时，设置UI状态
        /// </summary>
        private void SetUIForNoWorkbooksState()
        {
            cmbShippingWorkbook.Items.Clear();
            cmbBillWorkbook.Items.Clear();
            cmbShippingSheet.Items.Clear();
            cmbBillSheet.Items.Clear();

            // 禁用大部分控件
            gbShipping.Enabled = false;
            gbBill.Enabled = false;
            btnStart.Enabled = false;
            
            lblStatus.Text = "未检测到打开的Excel/WPS文件。请打开文件后点击“刷新列表”。";
        }

        /// <summary>
        /// 当成功加载工作簿后，恢复UI状态
        /// </summary>
        private void RestoreUIState()
        {
            gbShipping.Enabled = true;
            gbBill.Enabled = true;
            btnStart.Enabled = true;
        }


        private void cmbShippingWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbShippingWorkbook, cmbShippingSheet);
        }

        private void cmbBillWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbBillWorkbook, cmbBillSheet);
        }

        private void LoadSheetsForWorkbook(ComboBox workbookCombo, ComboBox sheetCombo)
        {
            try
            {
                if (workbooks == null || workbooks.Count == 0 || workbookCombo.SelectedIndex < 0)
                {
                    sheetCombo.Items.Clear();
                    return;
                }

                WorkbookInfo selectedWorkbookInfo = workbooks[workbookCombo.SelectedIndex];
                Excel.Workbook selectedWorkbook = selectedWorkbookInfo.Workbook;
                sheetCombo.Items.Clear();
                
                List<string> sheetNames = ExcelAddin.GetWorksheetNames(selectedWorkbook);
                sheetCombo.Items.AddRange(sheetNames.ToArray());
                
                if (sheetCombo == cmbShippingSheet)
                {
                    SetDefaultSheet(sheetCombo, new string[] { "发货明细", "发货", "shipping", "ship" });
                }
                else if (sheetCombo == cmbBillSheet)
                {
                    SetDefaultSheet(sheetCombo, new string[] { "账单明细", "账单", "bill", "bills" });
                }
                
                lblStatus.Text = $"工作簿: {selectedWorkbook.Name} | 工作表: {sheetNames.Count} 个";
            }
            catch (Exception ex)
            {
                WriteLog("加载工作表失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show($"加载工作表失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                string itemLower = item.ToLower();
                foreach (string keyword in keywords)
                {
                    if (itemLower.Contains(keyword.ToLower()))
                    {
                        combo.SelectedItem = item;
                        return;
                    }
                }
            }
            
            if (combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// 新增的刷新按钮点击事件
        /// </summary>
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "正在刷新工作簿列表...";
            Application.DoEvents(); // Give UI feedback immediately
            LoadWorkbooks();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (isProcessing)
            {
                MessageBox.Show("正在处理中，请稍候...", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                if (!ValidateInput())
                    return;

                MultiWorkbookMatchConfig config = CreateMatchConfig();
                
                SetUIEnabled(false);
                isProcessing = true;
                progressBar.Visible = true;
                progressBar.Value = 0;
                lblStatus.Text = "正在初始化匹配任务...";

                backgroundWorker.RunWorkerAsync(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动匹配失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetUIEnabled(true);
                isProcessing = false;
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

        private MultiWorkbookMatchConfig CreateMatchConfig()
        {
            return new MultiWorkbookMatchConfig
            {
                ShippingWorkbook = workbooks[cmbShippingWorkbook.SelectedIndex].Workbook,
                BillWorkbook = workbooks[cmbBillWorkbook.SelectedIndex].Workbook,
                ShippingSheetName = cmbShippingSheet.SelectedItem.ToString(),
                BillSheetName = cmbBillSheet.SelectedItem.ToString(),
                ShippingTrackColumn = txtShippingTrackColumn.Text.Trim().ToUpper(),
                ShippingProductColumn = txtShippingProductColumn.Text.Trim().ToUpper(),
                ShippingNameColumn = txtShippingNameColumn.Text.Trim().ToUpper(),
                BillTrackColumn = txtBillTrackColumn.Text.Trim().ToUpper(),
                BillProductColumn = txtBillProductColumn.Text.Trim().ToUpper(),
                BillNameColumn = txtBillNameColumn.Text.Trim().ToUpper()
            };
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                MultiWorkbookMatchConfig config = e.Argument as MultiWorkbookMatchConfig;
                
                MatchService.ProgressReportDelegate progressCallback = (progress, message) =>
                {
                    backgroundWorker.ReportProgress(progress, message);
                };

                MatchService service = new MatchService();
                e.Result = service.ExecuteMatch(config, progressCallback);
            }
            catch (Exception ex)
            {
                e.Result = new MatchResult { Success = false, ErrorMessage = ex.Message };
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = Math.Min(e.ProgressPercentage, 100);
            if (e.UserState != null)
            {
                lblStatus.Text = e.UserState.ToString();
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SetUIEnabled(true);
            isProcessing = false;
            progressBar.Visible = false;

            if (e.Error != null)
            {
                MessageBox.Show($"处理过程中发生错误：{e.Error.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "处理出错！";
                return;
            }

            if (e.Result is MatchResult result)
            {
                if (!result.Success)
                {
                    MessageBox.Show($"匹配失败：{result.ErrorMessage}\n\n请查看日志获取详细信息。", "匹配失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    lblStatus.Text = "匹配失败！";
                }
                else if (result.MatchedCount == 0)
                {
                    MessageBox.Show($"匹配完成，但没有找到匹配的运单！\n\n处理的账单行数：{result.ProcessedRows}\n处理耗时：{result.ElapsedSeconds:F2} 秒", "未找到匹配项", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    lblStatus.Text = "完成，但未找到匹配。";
                }
                else
                {
                    ShowResult(result);
                }
            }
        }

        private void SetUIEnabled(bool enabled)
        {
            gbShipping.Enabled = enabled;
            gbBill.Enabled = enabled;
            btnRefresh.Enabled = enabled;
            btnStart.Enabled = enabled;
            btnSettings.Enabled = enabled;
            btnViewLogs.Enabled = enabled;
            
            if (enabled)
            {
                btnStart.Text = "🚀 开始匹配";
                btnStart.BackColor = Color.FromArgb(0, 123, 255);
            }
            else
            {
                btnStart.Text = "🔄 处理中...";
                btnStart.BackColor = Color.Gray;
            }
        }

        private void ShowResult(MatchResult result)
        {
            lblStatus.Text = $"🎉 任务完成！耗时 {result.ElapsedSeconds:F2} 秒";

            double rowsPerSecond = result.ProcessedRows > 0 && result.ElapsedSeconds > 0 ? result.ProcessedRows / result.ElapsedSeconds : 0;
            
            string summary = $"🎉 运单匹配任务完成！\n" +
                             $"================================\n\n" +
                             $"📊 处理统计：\n" +
                             $"  • 处理账单行数：{result.ProcessedRows:N0} 行\n" +
                             $"  • 成功匹配运单：{result.MatchedCount:N0} 个\n" +
                             $"  • 填充数据单元格：{result.UpdatedCells:N0} 个\n\n" +
                             $"⚡ 性能表现：\n" +
                             $"  • 总处理时间：{result.ElapsedSeconds:F2} 秒\n" +
                             $"  • 处理速度：{rowsPerSecond:F0} 行/秒\n\n" +
                             $"✅ 任务结果：\n" +
                             $"  • 数据已成功写入到账单明细表。";

            MessageBox.Show(summary, "任务完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSelectTrackCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingTrackColumn, "发货明细运单号列");
        private void btnSelectProductCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingProductColumn, "发货明细商品编码列");
        private void btnSelectNameCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingNameColumn, "发货明细商品名称列");
        private void btnSelectBillTrackCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbBillWorkbook, txtBillTrackColumn, "账单明细运单号列");
        private void btnSelectBillProductCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbBillWorkbook, txtBillProductColumn, "账单明细商品编码列");
        private void btnSelectBillNameCol_Click(object sender, EventArgs e) => SelectColumnForWorkbook(cmbBillWorkbook, txtBillNameColumn, "账单明细商品名称列");

        private void SelectColumnForWorkbook(ComboBox workbookCombo, TextBox targetTextBox, string title)
        {
            if (workbookCombo.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择工作簿！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

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
        
        private void btnSettings_Click(object sender, EventArgs e)
        {
            try
            {
                using (SettingsForm settingsForm = new SettingsForm())
                {
                    if (settingsForm.ShowDialog() == DialogResult.OK)
                    {
                        ApplySettings();
                        SetDefaultValues();
                        MessageBox.Show("设置已应用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开设置窗口失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnViewLogs_Click(object sender, EventArgs e)
        {
            try
            {
                string logPath = MatchService.GetLogFolderPath();
                if (System.IO.Directory.Exists(logPath))
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

        private void btnCancel_Click(object sender, EventArgs e) => this.Close();

        private void WriteLog(string message, LogLevel level) => MatchService.WriteLog($"[MatchForm] {message}", level);

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (isProcessing)
            {
                if (MessageBox.Show("正在处理中，确定要退出吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    backgroundWorker.CancelAsync();
                }
            }
            base.OnFormClosing(e);
        }
    }
}