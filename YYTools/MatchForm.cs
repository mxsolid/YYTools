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

        private Dictionary<string, List<ColumnInfo>> columnCache = new Dictionary<string, List<ColumnInfo>>();

        private Dictionary<ComboBox, string> comboBoxColumnTypeMap;

        // 统一管理UI侧的异步任务（如启动后的预解析）
        private AsyncTaskManager _uiTaskManager = new AsyncTaskManager();
        private System.Windows.Forms.Timer _debounceTimer;

        public MatchForm()
        {
            InitializeComponent();
            
            // 启用DPI感知
            DPIManager.EnableDpiAwarenessForAllControls(this);
            
            InitializeCustomComponents();
            InitializeBackgroundWorker();
            InitializeForm();
            
            // 记录窗体创建日志
            Logger.LogUserAction("主窗体创建", "MatchForm已初始化", "成功");
        }

        private void InitializeCustomComponents()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;
            
            // 添加窗体大小调整事件
            this.Resize += MatchForm_Resize;
            
            this.Shown += (s, e) =>
            {
                this.Activate();
                // 将耗时初始化延后到显示之后执行，保障"秒开"
                try
                {
                    this.BeginInvoke(new Action(() =>
                    {
                        try { LoadMatcherSettings(); } catch { }
                        try { StartInitialParsingIfNeeded(); } catch { }
                    }));
                }
                catch { }
            };
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

            // 初始化防抖定时器（避免输入未完成就触发重算造成卡顿）
            _debounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _debounceTimer.Tick += (s, e) => { _debounceTimer.Stop(); RefreshWritePreview(); };
            // 为分隔符输入增加防抖
            txtDelimiter.TextChanged += (s, e) => DebounceRefreshWritePreview();
            chkRemoveDuplicates.CheckedChanged += (s, e) => RefreshWritePreview();
            cmbSort.SelectedIndexChanged += (s, e) => RefreshWritePreview();
            
            cmbSort.SelectedIndex = 0;
        }
        
        private void MatchForm_Resize(object sender, EventArgs e)
        {
            try
            {
                // 确保所有面板能够正确跟随窗体大小变化
                int margin = 12;
                int availableWidth = this.ClientSize.Width - (margin * 2);
                
                // 调整发货明细配置面板
                if (gbShipping != null)
                {
                    gbShipping.Width = availableWidth;
                }
                
                // 调整账单明细配置面板
                if (gbBill != null)
                {
                    gbBill.Width = availableWidth;
                }
                
                // 调整任务配置面板
                if (gbOptions != null)
                {
                    gbOptions.Width = availableWidth;
                }
                
                // 调整写入预览面板
                if (gbWritePreview != null)
                {
                    gbWritePreview.Width = availableWidth;
                }
                
                // 调整按钮面板
                if (panelButtons != null)
                {
                    panelButtons.Width = this.ClientSize.Width;
                    // 重新定位按钮
                    btnClose.Left = panelButtons.Width - btnClose.Width - margin;
                    btnStart.Left = btnClose.Left - btnStart.Width - 10;
                }
                
                // 调整状态面板
                if (panelStatus != null)
                {
                    panelStatus.Width = this.ClientSize.Width;
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"窗体大小调整处理失败: {ex.Message}");
            }
        }

        private void DebounceRefreshWritePreview()
        {
            try
            {
                if (_debounceTimer != null)
                {
                    _debounceTimer.Stop();
                    _debounceTimer.Start();
                }
                else
                {
                    RefreshWritePreview();
                }
            }
            catch { }
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
                // 将耗时初始化延后到 Shown 阶段
                // 同时更新列解析最大并发
                try { DataManager.UpdateMaxConcurrency(settings.MaxThreads); } catch { }
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
            TryPrefetchOtherIfParallelPossible();
        }

        private void cmbBillSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateColumnComboBoxes(cmbBillWorkbook, cmbBillSheet, cmbBillTrackColumn, cmbBillProductColumn, cmbBillNameColumn);
            TryPrefetchOtherIfParallelPossible();
        }

        private void TryPrefetchOtherIfParallelPossible()
        {
            try
            {
                if (cmbShippingWorkbook.SelectedIndex < 0 || cmbBillWorkbook.SelectedIndex < 0) return;
                if (cmbShippingSheet.SelectedIndex < 0 || cmbBillSheet.SelectedIndex < 0) return;

                var shipWb = workbooks[cmbShippingWorkbook.SelectedIndex].Workbook;
                var billWb = workbooks[cmbBillWorkbook.SelectedIndex].Workbook;
                var shipSheetName = cmbShippingSheet.SelectedItem?.ToString();
                var billSheetName = cmbBillSheet.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(shipSheetName) || string.IsNullOrEmpty(billSheetName)) return;

                // 如果是同一工作簿同一工作表，无需并行
                bool sameTarget = string.Equals((shipWb.FullName ?? shipWb.Name), (billWb.FullName ?? billWb.Name), StringComparison.OrdinalIgnoreCase)
                                  && string.Equals(shipSheetName, billSheetName, StringComparison.OrdinalIgnoreCase);
                if (sameTarget) return;

                // 使用线程池统一调度管理，避免创建过多线程
                int maxThreads = Math.Max(1, Math.Min(AppSettings.Instance.MaxThreads, Environment.ProcessorCount));
                
                // 使用异步任务管理器，避免阻塞UI
                _uiTaskManager.StartBackgroundTask(
                    taskName: "ParallelPrefetchColumns",
                    taskFactory: async (token, progress) =>
                    {
                        try
                        {
                            progress?.Report(new TaskProgress(10, "正在并行预取列信息..."));
                            
                            var tasks = new List<System.Threading.Tasks.Task>();
                            var semaphore = new System.Threading.SemaphoreSlim(maxThreads);
                            
                            // 并行预取两个工作表的列信息
                            Action<Excel.Workbook, string> prefetch = (wb, sheetName) =>
                            {
                                tasks.Add(System.Threading.Tasks.Task.Run(async () =>
                                {
                                    await semaphore.WaitAsync();
                                    try
                                    {
                                        var ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                                        if (ws != null)
                                        {
                                            // 触发一次列信息解析并写入缓存（如果已缓存则瞬间返回）
                                            var cols = DataManager.GetColumnInfos(ws);
                                            Logger.LogInfo($"并行预取列信息: {wb.Name}/{sheetName} 列数: {cols?.Count ?? 0}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogWarning($"并行预取失败: {wb.Name}/{sheetName} - {ex.Message}");
                                    }
                                    finally { semaphore.Release(); }
                                }));
                            };

                            // 对两个目标进行并行预取
                            prefetch(shipWb, shipSheetName);
                            prefetch(billWb, billSheetName);

                            progress?.Report(new TaskProgress(50, "等待预取完成..."));
                            await System.Threading.Tasks.Task.WhenAll(tasks);
                            
                            progress?.Report(new TaskProgress(100, "列信息预取完成"));
                            Logger.LogInfo("并行预取列信息完成");
                        }
                        catch (Exception ex)
                        {
                            Logger.LogWarning($"并行预取初始化失败: {ex.Message}");
                        }
                    },
                    allowMultiple: false
                );
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"并行预取初始化失败: {ex.Message}");
            }
        }

        private void LoadSheetsForWorkbook(ComboBox workbookCombo, ComboBox sheetCombo)
        {
            // 确保在UI线程中执行UI操作
            if (IsHandleCreated && !IsDisposed)
            {
                BeginInvoke(new Action(() =>
                {
                    sheetCombo.Items.Clear();
                    toolTip1.SetToolTip(sheetCombo, "");
                }));
            }
            
            if (workbookCombo.SelectedIndex < 0 || workbookCombo.SelectedIndex >= workbooks.Count) return;

            try
            {
                Excel.Workbook selectedWorkbook = workbooks[workbookCombo.SelectedIndex].Workbook;
                // 使用 DataManager 缓存工作表列表
                List<string> sheetNames = DataManager.GetSheetNames(selectedWorkbook);
                
                // 在UI线程中更新工作表列表
                if (IsHandleCreated && !IsDisposed)
                {
                    BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            sheetCombo.Items.AddRange(sheetNames.ToArray());
                            toolTip1.SetToolTip(sheetCombo, $"在工作簿 '{selectedWorkbook.Name}' 中选择一个工作表");

                            string[] keywords = sheetCombo == cmbShippingSheet ? new[] { "发货明细", "发货" } : new[] { "账单明细", "账单" };
                            SetDefaultSheet(sheetCombo, keywords);
                        }
                        catch (Exception ex)
                        {
                            WriteLog("加载工作表失败: " + ex.Message, LogLevel.Error);
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                WriteLog("加载工作表失败: " + ex.Message, LogLevel.Error);
            }
        }

        private void PopulateColumnComboBoxes(ComboBox wbCombo, ComboBox wsCombo, params ComboBox[] columnCombos)
        {
            // 确保在UI线程中执行初始清理操作
            if (IsHandleCreated && !IsDisposed)
            {
                BeginInvoke(new Action(() =>
                {
                    foreach (var combo in columnCombos) 
                    { 
                        combo.DataSource = null; 
                        combo.Items.Clear(); 
                        combo.Text = ""; 
                    }
                    toolTip1.SetToolTip(wsCombo, "请选择工作表");
                }));
            }

            if (wbCombo.SelectedIndex < 0 || wsCombo.SelectedIndex < 0 || wsCombo.SelectedItem == null) return;

            try
            {
                var wbInfo = workbooks[wbCombo.SelectedIndex];
                var ws = wbInfo.Workbook.Worksheets[wsCombo.SelectedItem.ToString()] as Excel.Worksheet;
                if (ws == null) return;

                // 在UI线程中显示加载状态
                if (IsHandleCreated && !IsDisposed)
                {
                    BeginInvoke(new Action(() => ShowLoading(true)));
                }
                
                // 使用异步任务管理器并行处理列信息
                _uiTaskManager.StartBackgroundTask(
                    taskName: "PopulateColumns",
                    taskFactory: async (token, progress) =>
                    {
                        try
                        {
                            progress?.Report(new TaskProgress(10, "正在解析列信息..."));
                            
                            // 优先从统一的 DataManager 缓存获取列信息
                            var columns = DataManager.GetColumnInfos(ws);
                            var cacheKey = $"{wbInfo.Name}_{wsCombo.SelectedItem}";
                            
                            // 线程安全地更新缓存
                            lock (columnCache)
                            {
                                columnCache[cacheKey] = columns;
                            }

                            progress?.Report(new TaskProgress(50, "正在获取工作表统计信息..."));
                            
                            // 并行获取工作表统计信息
                            var statsTask = System.Threading.Tasks.Task.Run(() =>
                            {
                                try
                                {
                                    var stats = ExcelHelper.GetWorksheetStats(ws);
                                    return new { rows = stats.rows, columns = stats.columns };
                                }
                                catch
                                {
                                    return new { rows = 0, columns = 0 };
                                }
                            });

                            // 并行处理智能列匹配
                            var smartMatchTask = System.Threading.Tasks.Task.Run(() =>
                            {
                                try
                                {
                                    if (settings.EnableSmartColumnSelection)
                                    {
                                        return SmartColumnService.SmartMatchColumns(columns);
                                    }
                                    return new Dictionary<string, ColumnInfo>();
                                }
                                catch
                                {
                                    return new Dictionary<string, ColumnInfo>();
                                }
                            });

                            // 等待并行任务完成
                            await System.Threading.Tasks.Task.WhenAll(statsTask, smartMatchTask);
                            
                            var worksheetStats = statsTask.Result;
                            var matchedColumns = smartMatchTask.Result;

                            progress?.Report(new TaskProgress(80, "正在更新UI..."));
                            
                            // 在UI线程中更新界面
                            if (IsHandleCreated && !IsDisposed)
                            {
                                BeginInvoke(new Action(() =>
                                {
                                    try
                                    {
                                        // 更新统计信息提示
                                        if (worksheetStats.rows > 0 || worksheetStats.columns > 0)
                                        {
                                            string statsString = $"总行数: {worksheetStats.rows:N0} | 总列数: {worksheetStats.columns:N0}";
                                            toolTip1.SetToolTip(wsCombo, statsString);
                                        }

                                        // 更新列下拉框
                                        foreach (var combo in columnCombos)
                                        {
                                            combo.DisplayMember = "ToString";
                                            combo.ValueMember = "ColumnLetter";
                                            combo.DataSource = new BindingSource(columns, null);
                                            combo.SelectedIndex = -1;
                                        }

                                        // 应用智能列选择
                                        if (matchedColumns.Count > 0)
                                        {
                                            ApplySmartColumnSelection(columnCombos, matchedColumns);
                                            
                                            foreach (var combo in columnCombos)
                                            {
                                                if (combo.SelectedItem != null && comboBoxColumnTypeMap.ContainsKey(combo))
                                                {
                                                    ValidateAndUpdateColumnInfo(combo);
                                                }
                                            }
                                        }

                                        // 更新状态和预览
                                        lblStatus.Text = $"已加载 {workbooks.Count} 个工作簿。请配置并开始任务。";
                                        RefreshWritePreview();
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogError($"更新列下拉框UI失败: {ex.Message}");
                                    }
                                }));
                            }
                            
                            progress?.Report(new TaskProgress(100, "列信息加载完成"));
                            Logger.LogInfo($"列信息加载完成: {wbInfo.Name}/{wsCombo.SelectedItem} 列数: {columns?.Count ?? 0}");
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError($"并行处理列信息失败: {ex.Message}");
                        }
                        finally
                        {
                            // 在UI线程中隐藏加载状态
                            if (IsHandleCreated && !IsDisposed)
                            {
                                BeginInvoke(new Action(() => ShowLoading(false)));
                            }
                        }
                    },
                    allowMultiple: false
                );
            }
            catch (Exception ex)
            {
                WriteLog("填充列下拉框失败: " + ex.Message, LogLevel.Error);
                // 在UI线程中隐藏加载状态
                if (IsHandleCreated && !IsDisposed)
                {
                    BeginInvoke(new Action(() => ShowLoading(false)));
                }
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
            
            // 第一优先级：完全匹配
            foreach (string item in combo.Items)
            {
                if (keywords.Any(k => string.Equals(item, k, StringComparison.OrdinalIgnoreCase)))
                {
                    combo.SelectedItem = item;
                    return;
                }
            }
            
            // 第二优先级：包含关键字的模糊匹配
            foreach (string item in combo.Items)
            {
                if (keywords.Any(k => item.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    combo.SelectedItem = item;
                    return;
                }
            }
            
            // 第三优先级：选择第一个可用项
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
        private void refreshListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Logger.LogUserAction("点击刷新列表", "清空缓存并重新加载工作簿", "开始");
                DataManager.ClearCache();
                columnCache.Clear();
                RefreshWorkbookList();
                Logger.LogUserAction("刷新列表完成", "", "成功");
            }
            catch (Exception ex)
            {
                Logger.LogError("刷新列表失败", ex);
                MessageBox.Show($"刷新失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
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
                TaskOptionsForm.ShowTaskOptions(this);
                // 重新加载设置
                LoadMatcherSettings();
                RefreshWritePreview();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开任务选项窗口失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try { _uiTaskManager?.Dispose(); } catch { }
            base.OnFormClosed(e);
        }

        /// <summary>
        /// 启动后进行轻量预解析：
        /// 1) 若无打开的工作簿，引导用户选择文件；
        /// 2) 获取活动工作簿与工作表，预热列信息缓存；
        /// 3) 全程异步，并在状态栏与进度条显示进度。
        /// </summary>
        private void StartInitialParsingIfNeeded()
        {
            try
            {
                excelApp = ExcelAddin.GetExcelApplication();
                if (excelApp == null || !ExcelAddin.HasOpenWorkbooks(excelApp))
                {
                    Logger.LogInfo("未检测到打开的Excel/WPS文件，弹出文件选择器");
                    openFileToolStripMenuItem_Click(this, EventArgs.Empty);
                }

                // 刷新一次工作簿列表，确保界面先可用
                RefreshWorkbookList();

                // 后台预解析当前活动工作簿
                _uiTaskManager.StartBackgroundTask(
                    taskName: "InitialParseActiveWorkbook",
                    taskFactory: async (token, progress) =>
                    {
                        try
                        {
                            await System.Threading.Tasks.Task.Yield();
                            var app = ExcelAddin.GetExcelApplication();
                            if (app == null || !ExcelAddin.HasOpenWorkbooks(app)) return;

                            Excel.Workbook activeWb = null;
                            try { activeWb = app.ActiveWorkbook; } catch { }
                            if (activeWb == null) return;

                            var wbName = activeWb.Name;
                            progress?.Report(new TaskProgress(10, $"正在预解析: {wbName}"));
                            Logger.LogUserAction("启动后预解析", $"活动工作簿: {wbName}", "开始");

                            // 预热工作表列表
                            var sheetNames = DataManager.GetSheetNames(activeWb);

                            // 获取活动工作表
                            Excel.Worksheet activeSheet = null;
                            try { activeSheet = app.ActiveSheet as Excel.Worksheet; } catch { }
                            if (activeSheet != null)
                            {
                                progress?.Report(new TaskProgress(40, $"解析工作表列: {activeSheet.Name}"));
                                var columns = DataManager.GetColumnInfos(activeSheet);
                                Logger.LogInfo($"预解析完成: {wbName}/{activeSheet.Name} 列数: {columns?.Count ?? 0}");
                            }

                            progress?.Report(new TaskProgress(100, "预解析完成"));
                            Logger.LogUserAction("启动后预解析", $"活动工作簿: {wbName}", "成功");
                        }
                        catch (System.OperationCanceledException)
                        {
                            Logger.LogUserAction("启动后预解析", "", "已取消");
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError("启动后预解析失败", ex);
                        }
                        finally
                        {
                            // UI提示
                            try
                            {
                                if (this.IsHandleCreated)
                                {
                                    this.BeginInvoke(new Action(() =>
                                    {
                                        progressBar.Style = ProgressBarStyle.Blocks;
                                        lblStatus.Text = "已加载工作簿列表。";
                                        progressBar.Visible = false;
                                    }));
                                }
                            }
                            catch { }
                        }
                    },
                    allowMultiple: false
                );

                // UI层展示加载动画
                ShowLoading(true);
                lblStatus.Text = "正在预解析活动工作簿...";
            }
            catch (Exception ex)
            {
                Logger.LogError("启动预解析初始化失败", ex);
            }
        }

        // 主界面不再提供“任务选项”入口，统一在菜单 工具->设置 中维护
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
            string aboutInfo = $"YY 运单匹配工具 {Constants.AppVersion}\n" +
                             $"版本哈希: {Constants.AppVersionHash}\n\n" +
                             "功能特点：\n" +
                             "• 智能运单匹配，支持灵活拼接\n" +
                             "• 优化智能列算法，提高准确率\n" +
                             "• 支持多工作簿操作与动态加载\n" +
                             "• 高性能处理，支持大数据量\n" +
                             "• 优化写入预览，配置更直观\n" +
                             "• 多线程并行处理，提升性能\n" +
                             "• 写入预览行数可配置\n\n" +
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
                // 确保在UI线程中执行所有UI操作
                if (!IsHandleCreated || IsDisposed) return;
                
                txtWritePreview.Text = "";
                if (cmbShippingWorkbook.SelectedIndex < 0 || cmbShippingSheet.SelectedIndex < 0 || cmbShippingSheet.SelectedItem == null)
                {
                    txtWritePreview.Text = "请先选择发货明细...";
                    return;
                }

                string trackCol = GetSelectedColumn(cmbShippingTrackColumn);
                string prodCol = GetSelectedColumn(cmbShippingProductColumn);
                string nameCol = GetSelectedColumn(cmbShippingNameColumn);

                if (string.IsNullOrEmpty(trackCol) || !ExcelHelper.IsValidColumnLetter(trackCol))
                {
                    txtWritePreview.Text = "请先选择有效的\"发货\"运单号列。";
                    return;
                }
                if (string.IsNullOrEmpty(prodCol) && string.IsNullOrEmpty(nameCol))
                {
                    txtWritePreview.Text = "请选择\"商品编码\"或\"商品名称\"列以生成预览。";
                    return;
                }

                var wbInfo = workbooks[cmbShippingWorkbook.SelectedIndex];
                var ws = wbInfo.Workbook.Worksheets[cmbShippingSheet.SelectedItem.ToString()] as Excel.Worksheet;
                if (ws == null) return;

                // 使用异步任务管理器在后台线程中处理Excel数据
                _uiTaskManager.StartBackgroundTask(
                    taskName: "RefreshWritePreview",
                    taskFactory: async (token, progress) =>
                    {
                        try
                        {
                            progress?.Report(new TaskProgress(10, "正在解析预览数据..."));
                            
                            Dictionary<string, List<ShippingItem>> previewIndex = new Dictionary<string, List<ShippingItem>>();
                            // 使用配置的预览行数，默认为20行
                            int maxScanRows = Math.Min(settings.PreviewParseRows, ws.UsedRange.Rows.Count);
                            int trackColNum = ExcelHelper.GetColumnNumber(trackCol);
                            int prodColNum = !string.IsNullOrEmpty(prodCol) && ExcelHelper.IsValidColumnLetter(prodCol) ? ExcelHelper.GetColumnNumber(prodCol) : -1;
                            int nameColNum = !string.IsNullOrEmpty(nameCol) && ExcelHelper.IsValidColumnLetter(nameCol) ? ExcelHelper.GetColumnNumber(nameCol) : -1;

                            progress?.Report(new TaskProgress(30, "正在并行解析数据..."));

                            // 多线程并行解析预览数据
                            var tasks = new List<System.Threading.Tasks.Task>();
                            var semaphore = new System.Threading.SemaphoreSlim(Math.Min(4, Environment.ProcessorCount));
                            
                            // 分批处理，每批处理一部分行
                            int batchSize = Math.Max(1, maxScanRows / 4);
                            for (int batchStart = 2; batchStart <= maxScanRows; batchStart += batchSize)
                            {
                                int batchEnd = Math.Min(batchStart + batchSize - 1, maxScanRows);
                                int startRow = batchStart;
                                int endRow = batchEnd;
                                
                                tasks.Add(System.Threading.Tasks.Task.Run(async () =>
                                {
                                    await semaphore.WaitAsync();
                                    try
                                    {
                                        var batchIndex = new Dictionary<string, List<ShippingItem>>();
                                        
                                        for (int r = startRow; r <= endRow; r++)
                                        {
                                            try
                                            {
                                                string trackNumber = ExcelHelper.GetCellValue(ws.Cells[r, trackColNum]);
                                                if (string.IsNullOrWhiteSpace(trackNumber)) continue;

                                                if (!batchIndex.ContainsKey(trackNumber))
                                                {
                                                    batchIndex[trackNumber] = new List<ShippingItem>();
                                                }

                                                batchIndex[trackNumber].Add(new ShippingItem
                                                {
                                                    ProductCode = prodColNum > 0 ? ExcelHelper.GetCellValue(ws.Cells[r, prodColNum]) : "",
                                                    ProductName = nameColNum > 0 ? ExcelHelper.GetCellValue(ws.Cells[r, nameColNum]) : ""
                                                });
                                            }
                                            catch (Exception ex)
                                            {
                                                Logger.LogWarning($"预览解析行 {r} 失败: {ex.Message}");
                                            }
                                        }
                                        
                                        // 线程安全地合并结果
                                        lock (previewIndex)
                                        {
                                            foreach (var kvp in batchIndex)
                                            {
                                                if (!previewIndex.ContainsKey(kvp.Key))
                                                {
                                                    previewIndex[kvp.Key] = new List<ShippingItem>();
                                                }
                                                previewIndex[kvp.Key].AddRange(kvp.Value);
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        semaphore.Release();
                                    }
                                }));
                            }

                            progress?.Report(new TaskProgress(70, "等待解析完成..."));

                            // 等待所有任务完成
                            await System.Threading.Tasks.Task.WhenAll(tasks);
                            
                            progress?.Report(new TaskProgress(90, "正在生成预览..."));
                            
                            var exampleEntry = previewIndex.FirstOrDefault(kvp => kvp.Value.Count > 1);
                            if (exampleEntry.Key == null) exampleEntry = previewIndex.FirstOrDefault();
                            if (exampleEntry.Key == null)
                            {
                                // 在UI线程中更新预览文本
                                if (IsHandleCreated && !IsDisposed)
                                {
                                    BeginInvoke(new Action(() =>
                                    {
                                        txtWritePreview.Text = $"（在前{maxScanRows}行发货明细中未找到可预览的数据）";
                                    }));
                                }
                                return;
                            }

                            List<string> previewLines = new List<string>();
                            List<ShippingItem> items = exampleEntry.Value;

                            if (prodColNum > 0)
                            {
                                IEnumerable<string> productCodes = items.Select(i => i.ProductCode).Where(pc => !string.IsNullOrWhiteSpace(pc));
                                if (productCodes.Any())
                                {
                                    previewLines.Add(BuildPreviewLine(productCodes, "商品: "));
                                }
                            }

                            if (nameColNum > 0)
                            {
                                IEnumerable<string> productNames = items.Select(i => i.ProductName).Where(pn => !string.IsNullOrWhiteSpace(pn));
                                if (productNames.Any())
                                {
                                    previewLines.Add(BuildPreviewLine(productNames, "品名: "));
                                }
                            }

                            string previewText = previewLines.Any() ? string.Join(Environment.NewLine, previewLines) : "（无有效数据可供预览）";
                            
                            // 在UI线程中更新预览文本和提示
                            if (IsHandleCreated && !IsDisposed)
                            {
                                BeginInvoke(new Action(() =>
                                {
                                    txtWritePreview.Text = previewText;
                                    toolTip1.SetToolTip(txtWritePreview, $"根据\"发货明细\"中的数据和下方选项，模拟匹配成功后将写入的数据效果。预览解析了前{maxScanRows}行数据。");
                                }));
                            }
                            
                            progress?.Report(new TaskProgress(100, "预览生成完成"));
                        }
                        catch (Exception ex)
                        {
                            WriteLog($"[MatchForm] 刷新写入预览失败: {ex.Message}", LogLevel.Warning);
                            // 在UI线程中显示错误信息
                            if (IsHandleCreated && !IsDisposed)
                            {
                                BeginInvoke(new Action(() =>
                                {
                                    txtWritePreview.Text = "生成预览时出错。";
                                }));
                            }
                        }
                    },
                    allowMultiple: false
                );
            }
            catch (Exception ex)
            {
                WriteLog($"[MatchForm] 刷新写入预览失败: {ex.Message}", LogLevel.Warning);
                txtWritePreview.Text = "生成预览时出错。";
            }
        }
        
        private string BuildPreviewLine(IEnumerable<string> data, string prefix)
        {
            string delimiter = txtDelimiter.Text;
            bool removeDuplicates = chkRemoveDuplicates.Checked;
            SortOption sortOption = GetSortOption();
            
            IEnumerable<string> processedData = data;

            if (removeDuplicates) processedData = processedData.Distinct();
            if (sortOption == SortOption.Asc) processedData = processedData.OrderBy(x => x, StringComparer.Ordinal);
            else if (sortOption == SortOption.Desc) processedData = processedData.OrderByDescending(x => x, StringComparer.Ordinal);
            
            return prefix + string.Join(delimiter, processedData);
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