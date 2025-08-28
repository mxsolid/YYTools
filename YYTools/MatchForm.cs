using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading; // 支持ThreadPool和Thread
using System.Threading.Tasks; // 异步任务支持

namespace YYTools
{
    public partial class MatchForm : Form
    {
        private Excel.Application excelApp;
        private BackgroundWorker backgroundWorker;
        private List<WorkbookInfo> workbooks = new List<WorkbookInfo>();
        private bool isProcessing = false;
        private AppSettings settings;

        private Dictionary<string, List<ColumnInfo>> columnCache = new Dictionary<string, List<ColumnInfo>>();

        private Dictionary<ComboBox, string> comboBoxColumnTypeMap;

        /// <summary>
        /// 内存监控定时器
        /// </summary>
        private System.Windows.Forms.Timer _memoryMonitorTimer;

        /// <summary>
        /// 任务超时保护
        /// </summary>
        private CancellationTokenSource _taskCancellationTokenSource;

        /// <summary>
        /// 初始化内存监控
        /// </summary>
        private void InitializeMemoryMonitor()
        {
            try
            {
                _memoryMonitorTimer = new System.Windows.Forms.Timer();
                _memoryMonitorTimer.Interval = 30000; // 30秒检查一次
                _memoryMonitorTimer.Tick += (s, e) =>
                {
                    try
                    {
                        var process = System.Diagnostics.Process.GetCurrentProcess();
                        var memoryMB = process.WorkingSet64 / 1024 / 1024;
                        
                        // 如果内存使用超过500MB，进行垃圾回收
                        if (memoryMB > 500)
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            
                            WriteLog($"内存使用: {memoryMB}MB，已执行垃圾回收", LogLevel.Info);
                        }
                        
                        // 如果内存使用超过1GB，显示警告
                        if (memoryMB > 1024)
                        {
                            WriteLog($"内存使用过高: {memoryMB}MB，建议重启程序", LogLevel.Warning);
                        }
                    }
                    catch
                    {
                        // 忽略内存监控错误
                    }
                };
                _memoryMonitorTimer.Start();
            }
            catch (Exception ex)
            {
                WriteLog($"初始化内存监控失败: {ex.Message}", LogLevel.Warning);
            }
        }

        public MatchForm()
        {
            InitializeComponent();
            
            // 基本初始化
            InitializeForm();
            
            // 延迟初始化其他组件，提高启动速度
            this.Load += (s, e) => DelayedInitialization();
        }

        /// <summary>
        /// 延迟初始化，提高启动速度
        /// </summary>
        private void DelayedInitialization()
        {
            try
            {
                // 在窗体加载完成后执行这些初始化
                ThreadPool.QueueUserWorkItem((state) =>
                {
                    try
                    {
                        // 延迟100ms执行，确保窗体完全显示
                        Thread.Sleep(100);
                        
                        if (this.InvokeRequired)
                        {
                            this.Invoke(new Action(() =>
                            {
                                try
                                {
                                    // DPI适配
                                    DPIManager.EnableDpiAwarenessForAllControls(this);
                                    
                                    // 其他初始化
                                    InitializeCustomComponents();
                                    InitializeBackgroundWorker();
                                    InitializeMemoryMonitor();
                                    
                                    // 异步加载Excel文件
                                    LoadExcelFiles();
                                    
                                    WriteLog("延迟初始化完成", LogLevel.Info);
                                }
                                catch (Exception ex)
                                {
                                    WriteLog($"延迟初始化失败: {ex.Message}", LogLevel.Warning);
                                }
                            }));
                        }
                    }
                    catch
                    {
                        // 忽略延迟初始化错误
                    }
                });
            }
            catch (Exception ex)
            {
                WriteLog($"启动延迟初始化失败: {ex.Message}", LogLevel.Warning);
            }
        }

        private void InitializeCustomComponents()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;
            this.Shown += (s, e) => { this.Activate(); };
            try
            {
                this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
                this.UpdateStyles();
            }
            catch { }

            comboBoxColumnTypeMap = new Dictionary<ComboBox, string>
            {
                { cmbShippingTrackColumn, "TrackColumn" },
                { cmbShippingProductColumn, "ProductColumn" },
                { cmbShippingNameColumn, "NameColumn" },
                { cmbBillTrackColumn, "TrackColumn" },
                { cmbBillProductColumn, "ProductColumn" },
                { cmbBillNameColumn, "NameColumn" }
            };

            txtDelimiter.TextChanged += (s, e) => RefreshWritePreview();
            chkRemoveDuplicates.CheckedChanged += (s, e) => RefreshWritePreview();
            cmbSort.SelectedIndexChanged += (s, e) => RefreshWritePreview();
            
            cmbSort.SelectedIndex = 0;
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
            toolTip1.SetToolTip(cmbShippingWorkbook, "选择包含发货明细的工作簿");
            toolTip1.SetToolTip(cmbBillWorkbook, "选择包含账单明细的工作簿");

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
            gbOptions.Enabled = false;
            gbWritePreview.Enabled = false;
            lblStatus.Text = "未检测到打开的Excel/WPS文件。请打开文件或从菜单栏选择文件。";
        }

        private void UpdateUIWithWorkbooks()
        {
            gbShipping.Enabled = true;
            gbBill.Enabled = true;
            btnStart.Enabled = true;
            gbOptions.Enabled = true;
            gbWritePreview.Enabled = true;
            lblStatus.Text = $"已加载 {workbooks.Count} 个工作簿。请配置并开始任务。";
        }

        private void cmbShippingWorkbook_SelectedIndexChanged(object sender, EventArgs e) => LoadSheetsForWorkbook(cmbShippingWorkbook, cmbShippingSheet);
        private void cmbBillWorkbook_SelectedIndexChanged(object sender, EventArgs e) => LoadSheetsForWorkbook(cmbBillWorkbook, cmbBillSheet);

        private void cmbShippingSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateColumnComboBoxes(cmbShippingWorkbook, cmbShippingSheet, cmbShippingTrackColumn, cmbShippingProductColumn, cmbShippingNameColumn);
        }

        private void cmbBillSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateColumnComboBoxes(cmbBillWorkbook, cmbBillSheet, cmbBillTrackColumn, cmbBillProductColumn, cmbBillNameColumn);
        }

        private void LoadSheetsForWorkbook(ComboBox workbookCombo, ComboBox sheetCombo)
        {
            sheetCombo.Items.Clear();
            toolTip1.SetToolTip(sheetCombo, "");
            if (workbookCombo.SelectedIndex < 0 || workbookCombo.SelectedIndex >= workbooks.Count) return;

            try
            {
                Excel.Workbook selectedWorkbook = workbooks[workbookCombo.SelectedIndex].Workbook;
                List<string> sheetNames = ExcelAddin.GetWorksheetNames(selectedWorkbook);
                sheetCombo.Items.AddRange(sheetNames.ToArray());
                toolTip1.SetToolTip(sheetCombo, $"在工作簿 '{selectedWorkbook.Name}' 中选择一个工作表");

                string[] keywords = sheetCombo == cmbShippingSheet ? new[] { "发货明细", "发货" } : new[] { "账单明细", "账单" };
                SetDefaultSheet(sheetCombo, keywords);
            }
            catch (Exception ex)
            {
                WriteLog("加载工作表失败: " + ex.Message, LogLevel.Error);
            }
        }

        private void PopulateColumnComboBoxes(ComboBox wbCombo, ComboBox wsCombo, params ComboBox[] columnCombos)
        {
            foreach (var combo in columnCombos) { combo.DataSource = null; combo.Items.Clear(); combo.Text = ""; }
            toolTip1.SetToolTip(wsCombo, "请选择工作表");

            if (wbCombo.SelectedIndex < 0 || wsCombo.SelectedIndex < 0 || wsCombo.SelectedItem == null) return;

            try
            {
                var wbInfo = workbooks[wbCombo.SelectedIndex];
                var ws = wbInfo.Workbook.Worksheets[wsCombo.SelectedItem.ToString()] as Excel.Worksheet;
                if (ws == null) return;

                ShowLoading(true);
                var columns = SmartColumnService.GetColumnInfos(ws, 50);
                var cacheKey = $"{wbInfo.Name}_{wsCombo.SelectedItem}";
                columnCache[cacheKey] = columns;

                try
                {
                    var stats = ExcelHelper.GetWorksheetStats(ws);
                    string statsString = $"总行数: {stats.rows:N0} | 总列数: {stats.columns:N0}";
                    toolTip1.SetToolTip(wsCombo, statsString);
                }
                catch { /* ignore */ }

                foreach (var combo in columnCombos)
                {
                    combo.DisplayMember = "ToString";
                    combo.ValueMember = "ColumnLetter";
                    combo.DataSource = new BindingSource(columns, null);
                    combo.SelectedIndex = -1;
                }

                if (settings.EnableSmartColumnSelection)
                {
                    var matchedColumns = SmartColumnService.SmartMatchColumns(columns);
                    ApplySmartColumnSelection(columnCombos, matchedColumns);
                    
                    foreach (var combo in columnCombos)
                    {
                        if (combo.SelectedItem != null && comboBoxColumnTypeMap.ContainsKey(combo))
                        {
                            ValidateAndUpdateColumnInfo(combo);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("填充列下拉框失败: " + ex.Message, LogLevel.Error);
            }
            finally
            {
                ShowLoading(false);
                lblStatus.Text = $"已加载 {workbooks.Count} 个工作簿。请配置并开始任务。";
                RefreshWritePreview();
            }
        }

        private void ShowLoading(bool loading)
        {
            try
            {
                progressBar.Visible = loading;
                progressBar.Style = loading ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
                if (loading)
                {
                    lblStatus.Text = "正在解析列信息...";
                }
            }
            catch { }
        }


        private void ApplySmartColumnSelection(ComboBox[] columnCombos, Dictionary<string, ColumnInfo> matchedColumns)
        {
            try
            {
                if (matchedColumns.ContainsKey("TrackColumn"))
                    SetColumnSelection(columnCombos[0], matchedColumns["TrackColumn"]);

                if (matchedColumns.ContainsKey("ProductColumn"))
                    SetColumnSelection(columnCombos[1], matchedColumns["ProductColumn"]);

                if (matchedColumns.ContainsKey("NameColumn"))
                    SetColumnSelection(columnCombos[2], matchedColumns["NameColumn"]);
            }
            catch (Exception ex)
            {
                WriteLog("应用智能列选择失败: " + ex.Message, LogLevel.Warning);
            }
        }

        private void SetColumnSelection(ComboBox combo, ColumnInfo columnInfo)
        {
            try
            {
                if (columnInfo != null)
                {
                    combo.SelectedValue = columnInfo.ColumnLetter;
                }
            }
            catch (Exception ex)
            {
                WriteLog($"设置列选择失败: {ex.Message}", LogLevel.Warning);
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
                settings.ConcatenationDelimiter = txtDelimiter.Text;
                settings.RemoveDuplicateItems = chkRemoveDuplicates.Checked;
                settings.Save();

                MultiWorkbookMatchConfig config = CreateMatchConfig();
                SetUiProcessingState(true);
                backgroundWorker.RunWorkerAsync(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动任务失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetUiProcessingState(false);
            }
        }

        private void SetUiProcessingState(bool processing)
        {
            this.isProcessing = processing;
            menuStrip1.Enabled = !processing;
            btnStart.Enabled = !processing;
            gbShipping.Enabled = !processing;
            gbBill.Enabled = !processing;
            gbOptions.Enabled = !processing;

            progressBar.Visible = processing;

            if (processing)
            {
                progressBar.Value = 0;
                lblStatus.Text = "正在初始化任务...";
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

                if (!AreColumnsValid(shippingWb.Workbook, cmbShippingSheet.Text, "发货", cmbShippingTrackColumn, cmbShippingProductColumn, cmbShippingNameColumn)) return false;
                if (!AreColumnsValid(billWb.Workbook, cmbBillSheet.Text, "账单", cmbBillTrackColumn, cmbBillProductColumn, cmbBillNameColumn)) return false;

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"验证列时发生错误: {ex.Message}", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private string GetSelectedColumn(ComboBox combo)
        {
            if (combo.SelectedValue == null) return combo.Text;
            return combo.SelectedValue.ToString().ToUpper();
        }

        private bool AreColumnsValid(Excel.Workbook wb, string sheetName, string type, params ComboBox[] columnCombos)
        {
            if (wb.Worksheets[sheetName] is Excel.Worksheet ws)
            {
                foreach (var cb in columnCombos)
                {
                    string colLetter = GetSelectedColumn(cb);
                    bool isValidFormat = !string.IsNullOrEmpty(colLetter) && ExcelHelper.IsValidColumnLetter(colLetter);

                    var cacheKey = $"{wb.Name}_{sheetName}";
                    bool existsInCache = columnCache.ContainsKey(cacheKey) &&
                                       columnCache[cacheKey].Any(col => col.ColumnLetter == colLetter);

                    if (!isValidFormat || !existsInCache)
                    {
                        MessageBox.Show($"您为“{type}”表选择的列“{cb.Text}”无效或不存在。", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                ShippingTrackColumn = GetSelectedColumn(cmbShippingTrackColumn),
                ShippingProductColumn = GetSelectedColumn(cmbShippingProductColumn),
                ShippingNameColumn = GetSelectedColumn(cmbShippingNameColumn),
                BillTrackColumn = GetSelectedColumn(cmbBillTrackColumn),
                BillProductColumn = GetSelectedColumn(cmbBillProductColumn),
                BillNameColumn = GetSelectedColumn(cmbBillNameColumn),
                SortOption = GetSortOption()
            };
        }

        private SortOption GetSortOption()
        {
            try
            {
                var text = cmbSort?.SelectedItem?.ToString() ?? "默认排序";
                if (text == "升序") return SortOption.Asc;
                if (text == "降序") return SortOption.Desc;
                return SortOption.None;
            }
            catch { return SortOption.None; }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var config = e.Argument as MultiWorkbookMatchConfig;
            var service = new MatchService();
            service.CancellationCheck = () => backgroundWorker.CancellationPending;

            e.Result = service.ExecuteMatch(config, (progress, message) =>
            {
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
                    lblStatus.Text = result.ErrorMessage == "任务被用户取消" ? "任务已由用户停止。" : "任务失败！";
                    if (result.ErrorMessage != "任务被用户取消")
                    {
                        MessageBox.Show($"任务失败：{result.ErrorMessage}\n\n请查看日志获取详细信息。", "任务失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
               ? $"任务完成，但没有找到匹配的运单！\n\n处理的账单行数：{result.ProcessedRows:N0}\n处理耗时：{result.ElapsedSeconds:F2} 秒"
               : $"🎉 任务完成！\n================================\n\n" +
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
                    if (settingsForm.ShowDialog(this) == DialogResult.OK)
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

        private void taskOptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // 显示任务选项配置窗体
                TaskOptionsForm.ShowTaskOptions(this);
                
                // 重新加载任务选项设置
                LoadMatcherSettings();
                
                // 刷新写入预览
                RefreshWritePreview();
                
                // 跳过Logger调用，避免Logger系统问题
                // Logger.LogUserAction("打开任务选项配置", "任务选项配置已更新", "成功");
            }
            catch (Exception ex)
            {
                // 跳过Logger调用，避免Logger系统问题
                // Logger.LogError("打开任务选项配置失败", ex);
                MessageBox.Show($"打开任务选项配置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            string aboutInfo = "YY 运单匹配工具 v2.10 (稳定修复版)\n\n" +
                             "功能特点：\n" +
                             "• 智能运单匹配，支持灵活拼接\n" +
                             "• 优化智能列算法，提高准确率\n" +
                             "• 支持多工作簿操作与动态加载\n" +
                             "• 高性能处理，支持大数据量\n" +
                             "• 优化写入预览，配置更直观\n\n" +
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

        private void cmbShippingTrackColumn_SelectedIndexChanged(object sender, EventArgs e) { ValidateAndUpdateColumnInfo(cmbShippingTrackColumn); RefreshWritePreview(); }
        private void cmbShippingProductColumn_SelectedIndexChanged(object sender, EventArgs e) { ValidateAndUpdateColumnInfo(cmbShippingProductColumn); RefreshWritePreview(); }
        private void cmbShippingNameColumn_SelectedIndexChanged(object sender, EventArgs e) { ValidateAndUpdateColumnInfo(cmbShippingNameColumn); RefreshWritePreview(); }
        private void cmbBillTrackColumn_SelectedIndexChanged(object sender, EventArgs e) => ValidateAndUpdateColumnInfo(cmbBillTrackColumn);
        private void cmbBillProductColumn_SelectedIndexChanged(object sender, EventArgs e) => ValidateAndUpdateColumnInfo(cmbBillProductColumn);
        private void cmbBillNameColumn_SelectedIndexChanged(object sender, EventArgs e) => ValidateAndUpdateColumnInfo(cmbBillNameColumn);

        private void ValidateAndUpdateColumnInfo(ComboBox combo)
        {
            try
            {
                toolTip1.SetToolTip(combo, combo.Text);
                combo.BackColor = SystemColors.Window;
            }
            catch (Exception ex)
            {
                WriteLog($"验证列信息失败: {ex.Message}", LogLevel.Warning);
            }
        }

        private void RefreshWritePreview()
        {
            try
            {
                // 显示进度条
                progressBar.Visible = true;
                progressBar.Value = 0;
                lblStatus.Text = "正在生成写入预览...";
                Application.DoEvents();

                txtWritePreview.Text = "";
                if (cmbShippingWorkbook.SelectedIndex < 0 || cmbShippingSheet.SelectedIndex < 0 || cmbShippingSheet.SelectedItem == null)
                {
                    txtWritePreview.Text = "请先选择发货明细...";
                    progressBar.Visible = false;
                    lblStatus.Text = "欢迎使用YY匹配工具";
                    return;
                }

                progressBar.Value = 20;
                Application.DoEvents();

                // 获取发货明细数据
                var shippingData = GetShippingData();
                if (shippingData == null || shippingData.Count == 0)
                {
                    txtWritePreview.Text = "未找到发货明细数据...";
                    progressBar.Visible = false;
                    lblStatus.Text = "欢迎使用YY匹配工具";
                    return;
                }

                progressBar.Value = 40;
                Application.DoEvents();

                // 获取账单明细数据
                var billData = GetBillData();
                if (billData == null || billData.Count == 0)
                {
                    txtWritePreview.Text = "未找到账单明细数据...";
                    progressBar.Visible = false;
                    lblStatus.Text = "欢迎使用YY匹配工具";
                    return;
                }

                progressBar.Value = 60;
                Application.DoEvents();

                // 生成预览数据
                var previewData = GeneratePreviewData(shippingData, billData);
                
                progressBar.Value = 80;
                Application.DoEvents();

                // 显示预览结果
                if (previewData.Count > 0)
                {
                    var previewText = string.Join("\n", previewData.Take(10)); // 只显示前10行
                    if (previewData.Count > 10)
                    {
                        previewText += $"\n... 还有 {previewData.Count - 10} 行数据";
                    }
                    txtWritePreview.Text = previewText;
                    lblStatus.Text = $"预览生成完成，共 {previewData.Count} 行数据";
                }
                else
                {
                    txtWritePreview.Text = "未找到匹配的数据...";
                    lblStatus.Text = "预览生成完成，无匹配数据";
                }

                progressBar.Value = 100;
                Application.DoEvents();

                // 延迟隐藏进度条 - 使用线程池提高安全性
                ThreadPool.QueueUserWorkItem((state) =>
                {
                    try
                    {
                        Thread.Sleep(1000); // 等待1秒
                        
                        if (this.InvokeRequired)
                        {
                            this.Invoke(new Action(() =>
                            {
                                try
                                {
                                    progressBar.Visible = false;
                                    progressBar.Value = 0;
                                }
                                catch
                                {
                                    // 忽略UI操作错误
                                }
                            }));
                        }
                        else
                        {
                            try
                            {
                                progressBar.Visible = false;
                                progressBar.Value = 0;
                            }
                            catch
                            {
                                // 忽略UI操作错误
                            }
                        }
                    }
                    catch
                    {
                        // 忽略线程池操作错误
                    }
                });
            }
            catch (Exception ex)
            {
                txtWritePreview.Text = $"生成预览失败: {ex.Message}";
                progressBar.Visible = false;
                lblStatus.Text = "预览生成失败";
            }
        }
        
        /// <summary>
        /// 构建预览行
        /// </summary>
        private string BuildPreviewLine(IEnumerable<string> items, string prefix)
        {
            try
            {
                if (items == null || !items.Any())
                    return $"{prefix}无数据";

                var delimiter = txtDelimiter.Text;
                if (string.IsNullOrWhiteSpace(delimiter))
                    delimiter = "、";

                var result = string.Join(delimiter, items.Where(i => !string.IsNullOrWhiteSpace(i)));
                return $"{prefix}{result}";
            }
            catch (Exception ex)
            {
                WriteLog($"构建预览行失败: {ex.Message}", LogLevel.Warning);
                return $"{prefix}构建失败";
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

        /// <summary>
        /// 获取发货明细数据
        /// </summary>
        private List<ShippingItem> GetShippingData()
        {
            try
            {
                if (cmbShippingWorkbook.SelectedIndex < 0 || cmbShippingSheet.SelectedIndex < 0)
                    return null;

                var wbInfo = workbooks[cmbShippingWorkbook.SelectedIndex];
                var ws = wbInfo.Workbook.Worksheets[cmbShippingSheet.SelectedItem.ToString()] as Excel.Worksheet;
                if (ws == null) return null;

                string trackCol = GetSelectedColumn(cmbShippingTrackColumn);
                string prodCol = GetSelectedColumn(cmbShippingProductColumn);
                string nameCol = GetSelectedColumn(cmbShippingNameColumn);

                if (string.IsNullOrEmpty(trackCol) || !ExcelHelper.IsValidColumnLetter(trackCol))
                    return null;

                var result = new List<ShippingItem>();
                int maxScanRows = Math.Min(100, ws.UsedRange.Rows.Count);
                int trackColNum = ExcelHelper.GetColumnNumber(trackCol);
                int prodColNum = !string.IsNullOrEmpty(prodCol) && ExcelHelper.IsValidColumnLetter(prodCol) ? ExcelHelper.GetColumnNumber(prodCol) : -1;
                int nameColNum = !string.IsNullOrEmpty(nameCol) && ExcelHelper.IsValidColumnLetter(nameCol) ? ExcelHelper.GetColumnNumber(nameCol) : -1;

                for (int r = 2; r <= maxScanRows; r++)
                {
                    string trackNumber = ExcelHelper.GetCellValue(ws.Cells[r, trackColNum]);
                    if (string.IsNullOrWhiteSpace(trackNumber)) continue;

                    result.Add(new ShippingItem
                    {
                        TrackNumber = trackNumber,
                        ProductCode = prodColNum > 0 ? ExcelHelper.GetCellValue(ws.Cells[r, prodColNum]) : "",
                        ProductName = nameColNum > 0 ? ExcelHelper.GetCellValue(ws.Cells[r, nameColNum]) : ""
                    });
                }

                return result;
            }
            catch (Exception ex)
            {
                WriteLog($"获取发货明细数据失败: {ex.Message}", LogLevel.Warning);
                return null;
            }
        }

        /// <summary>
        /// 获取账单明细数据
        /// </summary>
        private List<BillItem> GetBillData()
        {
            try
            {
                if (cmbBillWorkbook.SelectedIndex < 0 || cmbBillSheet.SelectedIndex < 0)
                    return null;

                var wbInfo = workbooks[cmbBillWorkbook.SelectedIndex];
                var ws = wbInfo.Workbook.Worksheets[cmbBillSheet.SelectedItem.ToString()] as Excel.Worksheet;
                if (ws == null) return null;

                string trackCol = GetSelectedColumn(cmbBillTrackColumn);
                if (string.IsNullOrEmpty(trackCol) || !ExcelHelper.IsValidColumnLetter(trackCol))
                    return null;

                var result = new List<BillItem>();
                int maxScanRows = Math.Min(100, ws.UsedRange.Rows.Count);
                int trackColNum = ExcelHelper.GetColumnNumber(trackCol);

                for (int r = 2; r <= maxScanRows; r++)
                {
                    string trackNumber = ExcelHelper.GetCellValue(ws.Cells[r, trackColNum]);
                    if (string.IsNullOrWhiteSpace(trackNumber)) continue;

                    result.Add(new BillItem
                    {
                        TrackNumber = trackNumber
                    });
                }

                return result;
            }
            catch (Exception ex)
            {
                WriteLog($"获取账单明细数据失败: {ex.Message}", LogLevel.Warning);
                return null;
            }
        }

        /// <summary>
        /// 生成预览数据
        /// </summary>
        private List<string> GeneratePreviewData(List<ShippingItem> shippingData, List<BillItem> billData)
        {
            try
            {
                var result = new List<string>();
                
                if (shippingData == null || billData == null || shippingData.Count == 0 || billData.Count == 0)
                {
                    result.Add("数据不足，无法生成预览");
                    return result;
                }

                // 创建发货数据的索引
                var shippingDict = shippingData
                    .Where(s => !string.IsNullOrWhiteSpace(s.TrackNumber))
                    .GroupBy(s => s.TrackNumber.Trim())
                    .ToDictionary(g => g.Key, g => g.ToList());

                WriteLog($"发货数据索引创建完成，共 {shippingDict.Count} 个运单号", LogLevel.Info);

                // 处理账单数据
                int processedCount = 0;
                foreach (var bill in billData.Take(20)) // 处理前20个账单
                {
                    if (string.IsNullOrWhiteSpace(bill.TrackNumber))
                        continue;

                    string trackNumber = bill.TrackNumber.Trim();
                    
                    if (shippingDict.ContainsKey(trackNumber))
                    {
                        var items = shippingDict[trackNumber];
                        var previewLines = new List<string>();

                        // 添加运单号
                        result.Add($"运单号: {trackNumber}");

                        // 处理商品编码
                        var productCodes = items
                            .Where(i => !string.IsNullOrWhiteSpace(i.ProductCode))
                            .Select(i => i.ProductCode.Trim())
                            .Where(pc => !string.IsNullOrWhiteSpace(pc))
                            .Distinct()
                            .ToList();

                        if (productCodes.Count > 0)
                        {
                            string productCodeLine = BuildPreviewLine(productCodes, "商品编码: ");
                            result.Add(productCodeLine);
                            WriteLog($"运单 {trackNumber} 找到 {productCodes.Count} 个商品编码", LogLevel.Debug);
                        }

                        // 处理商品名称
                        var productNames = items
                            .Where(i => !string.IsNullOrWhiteSpace(i.ProductName))
                            .Select(i => i.ProductName.Trim())
                            .Where(pn => !string.IsNullOrWhiteSpace(pn))
                            .Distinct()
                            .ToList();

                        if (productNames.Count > 0)
                        {
                            string productNameLine = BuildPreviewLine(productNames, "商品名称: ");
                            result.Add(productNameLine);
                            WriteLog($"运单 {trackNumber} 找到 {productNames.Count} 个商品名称", LogLevel.Debug);
                        }

                        if (previewLines.Count > 0)
                        {
                            result.AddRange(previewLines);
                        }

                        result.Add(""); // 空行分隔
                        processedCount++;
                    }
                    else
                    {
                        WriteLog($"运单号 {trackNumber} 在发货数据中未找到匹配", LogLevel.Debug);
                    }
                }

                if (processedCount == 0)
                {
                    result.Clear();
                    result.Add("未找到匹配的数据");
                    result.Add("可能的原因：");
                    result.Add("1. 运单号格式不一致（如空格、大小写等）");
                    result.Add("2. 发货明细和账单明细中的运单号不匹配");
                    result.Add("3. 数据为空或格式错误");
                    result.Add("");
                    result.Add($"发货数据: {shippingData.Count} 条");
                    result.Add($"账单数据: {billData.Count} 条");
                    result.Add($"发货运单号示例: {shippingData.Take(3).Select(s => s.TrackNumber).Where(t => !string.IsNullOrWhiteSpace(t)).FirstOrDefault() ?? "无"}");
                    result.Add($"账单运单号示例: {billData.Take(3).Select(b => b.TrackNumber).Where(t => !string.IsNullOrWhiteSpace(t)).FirstOrDefault() ?? "无"}");
                }
                else
                {
                    result.Insert(0, $"预览数据生成完成，共处理 {processedCount} 个运单");
                }

                WriteLog($"预览数据生成完成，共 {result.Count} 行", LogLevel.Info);
                return result;
            }
            catch (Exception ex)
            {
                WriteLog($"生成预览数据失败: {ex.Message}", LogLevel.Warning);
                return new List<string> { $"生成预览数据失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// 异步初始化Excel文件
        /// </summary>
        public async void InitializeExcelFilesAsync()
        {
            try
            {
                // 显示加载状态
                if (lblStatus.InvokeRequired)
                {
                    lblStatus.Invoke(new Action(() => lblStatus.Text = "正在加载Excel文件..."));
                }
                else
                {
                    lblStatus.Text = "正在加载Excel文件...";
                }

                // 异步加载Excel文件
                await Task.Run(() =>
                {
                    try
                    {
                        LoadExcelFiles();
                    }
                    catch (Exception ex)
                    {
                        // 记录错误但不影响程序运行
                        WriteLog($"异步加载Excel文件失败: {ex.Message}", LogLevel.Warning);
                    }
                });

                // 更新状态
                if (lblStatus.InvokeRequired)
                {
                    lblStatus.Invoke(new Action(() => lblStatus.Text = "Excel文件加载完成"));
                }
                else
                {
                    lblStatus.Text = "Excel文件加载完成";
                }

                // 延迟清除状态
                ThreadPool.QueueUserWorkItem((state) =>
                {
                    try
                    {
                        Thread.Sleep(2000); // 等待2秒
                        if (lblStatus.InvokeRequired)
                        {
                            lblStatus.Invoke(new Action(() => lblStatus.Text = "欢迎使用YY匹配工具"));
                        }
                        else
                        {
                            lblStatus.Text = "欢迎使用YY匹配工具";
                        }
                    }
                    catch
                    {
                        // 忽略错误
                    }
                });
            }
            catch (Exception ex)
            {
                WriteLog($"异步初始化Excel文件失败: {ex.Message}", LogLevel.Warning);
                if (lblStatus.InvokeRequired)
                {
                    lblStatus.Invoke(new Action(() => lblStatus.Text = "Excel文件加载失败"));
                }
                else
                {
                    lblStatus.Text = "Excel文件加载失败";
                }
            }
        }

        /// <summary>
        /// 开始任务（带超时保护）
        /// </summary>
        private async void StartTaskWithTimeout()
        {
            try
            {
                // 取消之前的任务
                _taskCancellationTokenSource?.Cancel();
                _taskCancellationTokenSource = new CancellationTokenSource();

                // 设置超时时间（5分钟）
                var timeoutToken = new CancellationTokenSource(TimeSpan.FromMinutes(5));
                var combinedToken = CancellationTokenSource.CreateLinkedTokenSource(
                    _taskCancellationTokenSource.Token, timeoutToken.Token);

                // 显示进度条
                progressBar.Visible = true;
                progressBar.Value = 0;
                lblStatus.Text = "正在执行任务...";
                btnStart.Enabled = false;
                Application.DoEvents();

                try
                {
                    // 执行任务
                    await Task.Run(() =>
                    {
                        // 这里执行实际的匹配任务
                        // 为了演示，我们模拟一个长时间运行的任务
                        for (int i = 0; i < 100; i++)
                        {
                            if (combinedToken.Token.IsCancellationRequested)
                                throw new OperationCanceledException();
                            
                            // 更新进度
                            var progress = i;
                            if (progressBar.InvokeRequired)
                            {
                                progressBar.Invoke(new Action(() => progressBar.Value = progress));
                            }
                            else
                            {
                                progressBar.Value = progress;
                            }
                            
                            Thread.Sleep(100); // 模拟工作
                        }
                    }, combinedToken.Token);

                    // 任务完成
                    lblStatus.Text = "任务执行完成";
                    progressBar.Value = 100;
                }
                catch (OperationCanceledException)
                {
                    // if (timeoutToken.Token.IsCancellationRequested())
                    // {
                    //     lblStatus.Text = "任务执行超时，已自动取消";
                    //     MessageBox.Show("任务执行时间过长，已自动取消。请检查数据量或优化配置。", "任务超时", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // }
                    // else
                    // {
                    //     lblStatus.Text = "任务已取消";
                    // }
                }
                catch (Exception ex)
                {
                    lblStatus.Text = $"任务执行失败: {ex.Message}";
                    MessageBox.Show($"任务执行失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // 恢复界面状态
                    btnStart.Enabled = true;
                    
                    // 延迟隐藏进度条
                    ThreadPool.QueueUserWorkItem((state) =>
                    {
                        try
                        {
                            Thread.Sleep(2000);
                            if (progressBar.InvokeRequired)
                            {
                                progressBar.Invoke(new Action(() =>
                                {
                                    progressBar.Visible = false;
                                    progressBar.Value = 0;
                                }));
                            }
                            else
                            {
                                progressBar.Visible = false;
                                progressBar.Value = 0;
                            }
                        }
                        catch
                        {
                            // 忽略错误
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                WriteLog($"启动任务失败: {ex.Message}", LogLevel.Error);
                lblStatus.Text = "启动任务失败";
                btnStart.Enabled = true;
                progressBar.Visible = false;
            }
        }

        /// <summary>
        /// 加载Excel文件列表
        /// </summary>
        private void LoadExcelFiles()
        {
            try
            {
                // 清空现有列表
                cmbShippingWorkbook.Items.Clear();
                cmbBillWorkbook.Items.Clear();
                workbooks.Clear();

                // 获取当前目录下的Excel文件
                string[] excelFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx")
                    .Concat(Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xls"))
                    .ToArray();

                if (excelFiles.Length == 0)
                {
                    WriteLog("当前目录下未找到Excel文件", LogLevel.Info);
                    return;
                }

                // 加载每个Excel文件
                foreach (string filePath in excelFiles)
                {
                    try
                    {
                        var fileInfo = new FileInfo(filePath);
                        if (fileInfo.Length > 100 * 1024 * 1024) // 100MB限制
                        {
                            WriteLog($"文件过大，跳过: {fileInfo.Name}", LogLevel.Warning);
                            continue;
                        }

                        var wbInfo = new WorkbookInfo
                        {
                            FilePath = filePath,
                            FileName = fileInfo.Name,
                            Name = fileInfo.Name, // 设置Name属性
                            FileSize = fileInfo.Length,
                            LastModified = fileInfo.LastWriteTime,
                            IsActive = false // 默认非活动状态
                        };

                        workbooks.Add(wbInfo);
                        cmbShippingWorkbook.Items.Add(fileInfo.Name);
                        cmbBillWorkbook.Items.Add(fileInfo.Name);

                        WriteLog($"已加载Excel文件: {fileInfo.Name} ({fileInfo.Length / 1024}KB)", LogLevel.Info);
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"加载Excel文件失败: {Path.GetFileName(filePath)}, 错误: {ex.Message}", LogLevel.Warning);
                    }
                }

                // 如果有文件，选择第一个
                if (cmbShippingWorkbook.Items.Count > 0)
                {
                    cmbShippingWorkbook.SelectedIndex = 0;
                    cmbBillWorkbook.SelectedIndex = 0;
                }

                WriteLog($"Excel文件加载完成，共 {workbooks.Count} 个文件", LogLevel.Info);
            }
            catch (Exception ex)
            {
                WriteLog($"加载Excel文件列表失败: {ex.Message}", LogLevel.Error);
            }
        }
    }
}