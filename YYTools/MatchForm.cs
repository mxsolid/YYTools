using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
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
            InitializeBackgroundWorker();
            
            // 彻底解决聚焦问题
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;
            this.TopMost = true;
            
            InitializeForm();
            
            // 确保窗体完全显示后再取消置顶
            this.Shown += (s, e) => 
            {
                this.TopMost = false;
                this.Activate();
                this.Focus();
                this.BringToFront();
            };
        }

        /// <summary>
        /// 初始化后台工作线程
        /// </summary>
        private void InitializeBackgroundWorker()
        {
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
        }

        /// <summary>
        /// 初始化窗体 - 简化版本
        /// </summary>
        private void InitializeForm()
        {
            try
            {
                // 应用设置
                ApplySettings();
                
                // 获取WPS/Excel应用程序实例
                excelApp = ExcelAddin.Application;
                
                // 检查连接
                if (excelApp == null)
                {
                    MessageBox.Show("请先打开WPS表格或Excel文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.Close();
                    return;
                }
                
                // 加载工作簿列表
                LoadWorkbooks();
                
                // 设置默认值
                SetDefaultValues();

                // 清理旧日志
                MatchService.CleanupOldLogs();

                // 设置焦点到第一个输入控件
                if (cmbShippingWorkbook.Items.Count > 0)
                {
                    cmbShippingWorkbook.Focus();
                }
            }
            catch (Exception ex)
            {
                WriteLog("初始化窗体失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show("初始化失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 应用设置到窗体
        /// </summary>
        private void ApplySettings()
        {
            try
            {
                AppSettings settings = AppSettings.Instance;
                
                // 应用字体设置
                Font newFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                ApplyFontToControls(this, newFont);
                
                // 应用DPI缩放
                if (settings.AutoScaleUI)
                {
                    this.AutoScaleMode = AutoScaleMode.Dpi;
                }
            }
            catch (Exception ex)
            {
                WriteLog("应用设置失败: " + ex.Message, LogLevel.Warning);
            }
        }

        /// <summary>
        /// 递归应用字体到所有控件
        /// </summary>
        private void ApplyFontToControls(Control parent, Font font)
        {
            try
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
            catch (Exception ex)
            {
                WriteLog("应用字体失败: " + ex.Message, LogLevel.Warning);
            }
        }

        /// <summary>
        /// 加载工作簿列表 - 增强版本
        /// </summary>
        private void LoadWorkbooks()
        {
            try
            {
                WriteLog("开始加载工作簿列表", LogLevel.Info);
                
                // 强制重新获取工作簿列表
                workbooks = ExcelAddin.GetOpenWorkbooks();
                
                cmbShippingWorkbook.Items.Clear();
                cmbBillWorkbook.Items.Clear();
                
                if (workbooks == null || workbooks.Count == 0)
                {
                    WriteLog("没有检测到打开的工作簿", LogLevel.Warning);
                    
                    // 再次尝试获取
                    System.Threading.Thread.Sleep(500);
                    workbooks = ExcelAddin.GetOpenWorkbooks();
                    
                    if (workbooks == null || workbooks.Count == 0)
                    {
                        MessageBox.Show("没有检测到打开的工作簿！\n\n调试信息：\n1. 请确保在WPS表格或Excel中已打开文件\n2. 文件不能是只读或受保护状态\n3. 尝试关闭工具重新打开", 
                            "检测失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.Close();
                        return;
                    }
                }
                
                WriteLog("检测到 " + workbooks.Count + " 个工作簿", LogLevel.Info);
                
                int activeIndex = -1;
                
                // 添加工作簿到下拉列表
                for (int i = 0; i < workbooks.Count; i++)
                {
                    var workbook = workbooks[i];
                    string displayName = workbook.Name;
                    
                    if (workbook.IsActive)
                    {
                        displayName += " [当前活动]";
                        activeIndex = i;
                        WriteLog("发现活动工作簿: " + workbook.Name, LogLevel.Info);
                    }
                    
                    cmbShippingWorkbook.Items.Add(displayName);
                    cmbBillWorkbook.Items.Add(displayName);
                    
                    WriteLog("添加工作簿: " + displayName, LogLevel.Info);
                }
                
                // 优先选择活动工作簿
                if (activeIndex >= 0)
                {
                    cmbShippingWorkbook.SelectedIndex = activeIndex;
                    cmbBillWorkbook.SelectedIndex = activeIndex;
                    WriteLog("自动选择活动工作簿: " + workbooks[activeIndex].Name, LogLevel.Info);
                }
                else if (workbooks.Count > 0)
                {
                    // 如果没有活动工作簿，选择第一个
                    cmbShippingWorkbook.SelectedIndex = 0;
                    cmbBillWorkbook.SelectedIndex = 0;
                    WriteLog("自动选择第一个工作簿: " + workbooks[0].Name, LogLevel.Info);
                }
                
                // 更新状态
                lblStatus.Text = string.Format("已加载 {0} 个工作簿{1}", 
                    workbooks.Count, 
                    activeIndex >= 0 ? "，已选择活动工作簿" : "");
                
                WriteLog("工作簿加载完成", LogLevel.Info);
            }
            catch (Exception ex)
            {
                WriteLog("加载工作簿失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show("加载工作簿失败：" + ex.Message + "\n\n请尝试：\n1. 重新启动WPS/Excel\n2. 确保文件正常打开\n3. 检查文件是否受保护", 
                    "加载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        /// <summary>
        /// 发货工作簿选择变化事件
        /// </summary>
        private void cmbShippingWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbShippingWorkbook, cmbShippingSheet);
        }

        /// <summary>
        /// 账单工作簿选择变化事件
        /// </summary>
        private void cmbBillWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheetsForWorkbook(cmbBillWorkbook, cmbBillSheet);
        }

        /// <summary>
        /// 为指定工作簿加载工作表 - 立即更新版
        /// </summary>
        private void LoadSheetsForWorkbook(ComboBox workbookCombo, ComboBox sheetCombo)
        {
            try
            {
                if (workbooks == null || workbooks.Count == 0)
                {
                    WriteLog("工作簿列表为空，尝试重新加载", LogLevel.Warning);
                    LoadWorkbooks();
                    return;
                }
                
                if (workbookCombo.SelectedIndex >= 0 && workbookCombo.SelectedIndex < workbooks.Count)
                {
                    WorkbookInfo selectedWorkbook = workbooks[workbookCombo.SelectedIndex];
                    sheetCombo.Items.Clear();
                    
                    List<string> sheetNames = ExcelAddin.GetWorksheetNames(selectedWorkbook.Workbook);
                    foreach (string sheetName in sheetNames)
                    {
                        sheetCombo.Items.Add(sheetName);
                    }
                    
                    // 智能自动选择工作表
                    if (sheetCombo == cmbShippingSheet)
                    {
                        SetDefaultSheet(sheetCombo, new string[] { "发货明细", "发货", "shipping", "ship" });
                    }
                    else if (sheetCombo == cmbBillSheet)
                    {
                        SetDefaultSheet(sheetCombo, new string[] { "账单明细", "账单", "bill", "bills" });
                    }
                    
                    // 立即刷新界面
                    sheetCombo.Refresh();
                    Application.DoEvents();
                    
                    // 更新状态信息
                    lblStatus.Text = string.Format("已选择工作簿: {0}，包含 {1} 个工作表", 
                        selectedWorkbook.Name, sheetNames.Count);
                }
                else
                {
                    // 清空工作表列表
                    sheetCombo.Items.Clear();
                    sheetCombo.Refresh();
                }
            }
            catch (Exception ex)
            {
                WriteLog("加载工作表失败: " + ex.Message, LogLevel.Error);
                MessageBox.Show(string.Format("加载工作表失败：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 设置默认值 - 从配置加载
        /// </summary>
        private void SetDefaultValues()
        {
            try
            {
                AppSettings settings = AppSettings.Instance;
                
                // 设置默认列 - 从配置文件加载
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
                
                // 如果加载配置失败，使用硬编码默认值
                txtShippingTrackColumn.Text = "B";
                txtShippingProductColumn.Text = "J";
                txtShippingNameColumn.Text = "I";
                
                txtBillTrackColumn.Text = "C";
                txtBillProductColumn.Text = "Y";
                txtBillNameColumn.Text = "Z";
            }
        }

        /// <summary>
        /// 根据关键字设置默认工作表 - 改进版
        /// </summary>
        private void SetDefaultSheet(ComboBox combo, string[] keywords, bool preferFirst = false)
        {
            if (combo.Items.Count == 0) return;

            // 首先尝试精确匹配关键字
            foreach (string item in combo.Items)
            {
                string itemLower = item.ToString().ToLower();
                foreach (string keyword in keywords)
                {
                    if (itemLower == keyword.ToLower() || itemLower.Contains(keyword.ToLower()))
                    {
                        combo.SelectedItem = item;
                        return;
                    }
                }
            }
            
            // 如果没有找到匹配的，根据preferFirst参数决定
            if (preferFirst && combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }
            else if (combo.Items.Count > 0)
            {
                // 默认选择第一个，但优先级较低
                combo.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// 开始匹配按钮点击事件
        /// </summary>
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (isProcessing)
            {
                MessageBox.Show("正在处理中，请稍候...", "提示", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 验证输入
                if (!ValidateInput())
                    return;

                // 创建匹配配置
                MultiWorkbookMatchConfig config = CreateMatchConfig();
                
                // 设置UI状态
                SetUIEnabled(false);
                isProcessing = true;
                progressBar.Visible = true;
                progressBar.Value = 0;
                lblStatus.Text = "正在初始化匹配任务...";

                // 启动后台匹配任务
                backgroundWorker.RunWorkerAsync(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("启动匹配失败：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetUIEnabled(true);
                isProcessing = false;
            }
        }

        /// <summary>
        /// 验证用户输入
        /// </summary>
        private bool ValidateInput()
        {
            if (cmbShippingWorkbook.SelectedIndex < 0)
            {
                MessageBox.Show("请选择发货明细工作簿！", "验证失败", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbShippingWorkbook.Focus();
                return false;
            }

            if (cmbBillWorkbook.SelectedIndex < 0)
            {
                MessageBox.Show("请选择账单明细工作簿！", "验证失败", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbBillWorkbook.Focus();
                return false;
            }

            if (cmbShippingSheet.SelectedIndex < 0)
            {
                MessageBox.Show("请选择发货明细工作表！", "验证失败", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbShippingSheet.Focus();
                return false;
            }

            if (cmbBillSheet.SelectedIndex < 0)
            {
                MessageBox.Show("请选择账单明细工作表！", "验证失败", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbBillSheet.Focus();
                return false;
            }

            return true;
        }

        /// <summary>
        /// 创建多工作簿匹配配置
        /// </summary>
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

        /// <summary>
        /// 后台工作线程 - 执行匹配
        /// </summary>
        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                MultiWorkbookMatchConfig config = e.Argument as MultiWorkbookMatchConfig;
                
                // 将多工作簿配置转换为标准配置
                MatchConfig standardConfig = new MatchConfig
                {
                    ShippingSheetName = config.ShippingSheetName,
                    BillSheetName = config.BillSheetName,
                    ShippingTrackColumn = config.ShippingTrackColumn,
                    ShippingProductColumn = config.ShippingProductColumn,
                    ShippingNameColumn = config.ShippingNameColumn,
                    BillTrackColumn = config.BillTrackColumn,
                    BillProductColumn = config.BillProductColumn,
                    BillNameColumn = config.BillNameColumn
                };

                // 创建临时Excel应用实例来处理跨工作簿操作
                Excel.Application tempApp = config.ShippingWorkbook.Application;

                MatchService.ProgressReportDelegate progressCallback = (progress, message) =>
                {
                    backgroundWorker.ReportProgress(progress, message);
                };

                MatchService service = new MatchService();
                MatchResult result = service.ExecuteMatch(standardConfig, tempApp, progressCallback);
                
                e.Result = result;
            }
            catch (Exception ex)
            {
                MatchResult errorResult = new MatchResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    ProcessedRows = 0,
                    MatchedCount = 0,
                    UpdatedCells = 0
                };
                e.Result = errorResult;
            }
        }

        /// <summary>
        /// 进度更新事件
        /// </summary>
        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                progressBar.Value = Math.Min(e.ProgressPercentage, 100);
                
                if (e.UserState != null)
                {
                    lblStatus.Text = e.UserState.ToString();
                }

                // 强制界面更新
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("更新进度时出错: " + ex.Message);
            }
        }

        /// <summary>
        /// 后台工作完成事件
        /// </summary>
        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                // 恢复界面状态
                SetUIEnabled(true);
                isProcessing = false;
                progressBar.Visible = false;
                lblStatus.Visible = false;

                if (e.Error != null)
                {
                    MessageBox.Show(string.Format("处理过程中发生错误：{0}", e.Error.Message), "错误", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                MatchResult result = e.Result as MatchResult;
                if (result != null)
                {
                    // 检查结果是否包含错误
                    if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
                    {
                        MessageBox.Show(string.Format("匹配失败：{0}\n\n请查看日志获取详细信息。", result.ErrorMessage), 
                            "匹配失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    
                    // 检查是否有匹配结果
                    if (result.Success && result.MatchedCount == 0)
                    {
                        MessageBox.Show(string.Format("匹配完成，但没有找到匹配的运单！\n\n处理的账单行数：{0}\n匹配的运单数：{1}\n处理耗时：{2:F2} 秒\n\n可能原因：\n1. 运单号格式不匹配\n2. 发货明细中没有对应的运单号\n3. 列设置不正确\n\n请检查数据或查看日志。", 
                            result.ProcessedRows, result.MatchedCount, result.ElapsedSeconds), 
                            "未找到匹配项", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    ShowResult(result);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("完成处理时发生错误：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 设置UI控件启用状态
        /// </summary>
        private void SetUIEnabled(bool enabled)
        {
            // 工作簿和工作表选择
            cmbShippingWorkbook.Enabled = enabled;
            cmbBillWorkbook.Enabled = enabled;
            cmbShippingSheet.Enabled = enabled;
            cmbBillSheet.Enabled = enabled;
            
            // 列设置文本框
            txtShippingTrackColumn.Enabled = enabled;
            txtShippingProductColumn.Enabled = enabled;
            txtShippingNameColumn.Enabled = enabled;
            txtBillTrackColumn.Enabled = enabled;
            txtBillProductColumn.Enabled = enabled;
            txtBillNameColumn.Enabled = enabled;
            
            // 所有选择列按钮
            btnSelectTrackCol.Enabled = enabled;
            btnSelectProductCol.Enabled = enabled;
            btnSelectNameCol.Enabled = enabled;
            btnSelectBillTrackCol.Enabled = enabled;
            btnSelectBillProductCol.Enabled = enabled;
            btnSelectBillNameCol.Enabled = enabled;
            
            // 主要操作按钮
            btnStart.Enabled = enabled;
            btnSettings.Enabled = enabled;
            btnViewLogs.Enabled = enabled;
            
            // 更新按钮文本和样式
            if (enabled)
            {
                btnStart.Text = "🚀 开始匹配";
                btnStart.BackColor = System.Drawing.Color.FromArgb(0, 123, 255);
            }
            else
            {
                btnStart.Text = "🔄 处理中...";
                btnStart.BackColor = System.Drawing.Color.Gray;
            }
        }

        /// <summary>
        /// 显示匹配结果
        /// </summary>
        private void ShowResult(MatchResult result)
        {
            if (result.Success)
            {
                string message = string.Format("匹配完成！\n\n处理的账单行数：{0}\n匹配的运单数：{1}\n填充的单元格数：{2}\n处理耗时：{3:F2} 秒",
                    result.ProcessedRows, result.MatchedCount, result.UpdatedCells, result.ElapsedSeconds);
                
                MessageBox.Show(message, "成功", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                string errorMsg = string.Format("匹配失败：{0}", result.ErrorMessage);
                if (result.ElapsedSeconds > 0)
                {
                    errorMsg += string.Format("\n耗时：{0:F2} 秒", result.ElapsedSeconds);
                }

                MessageBox.Show(errorMsg, "失败", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 选择列按钮点击事件
        /// </summary>
        private void btnSelectColumn_Click(object sender, EventArgs e)
        {
            Button clickedButton = sender as Button;
            TextBox targetTextBox = GetTargetTextBox(clickedButton);
            
            if (targetTextBox == null) return;

            try
            {
                // 显示选择对话框
                ColumnSelectionForm selectionForm = new ColumnSelectionForm(targetTextBox.Text);
                
                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    targetTextBox.Text = selectionForm.SelectedColumn;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("选择列时发生错误：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 根据按钮获取对应的文本框
        /// </summary>
        private TextBox GetTargetTextBox(Button button)
        {
            if (button == null) return null;

            switch (button.Name)
            {
                case "btnSelectShippingTrack":
                    return txtShippingTrackColumn;
                case "btnSelectShippingProduct":
                    return txtShippingProductColumn;
                case "btnSelectShippingName":
                    return txtShippingNameColumn;
                case "btnSelectBillTrack":
                    return txtBillTrackColumn;
                case "btnSelectBillProduct":
                    return txtBillProductColumn;
                case "btnSelectBillName":
                    return txtBillNameColumn;
                default:
                    return null;
            }
        }

        /// <summary>
        /// 取消按钮点击事件
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (isProcessing)
            {
                DialogResult result = MessageBox.Show("正在处理中，确定要退出吗？", "确认", 
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

        /// <summary>
        /// 查看日志按钮点击事件
        /// </summary>
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
                    MessageBox.Show("日志文件夹不存在，可能还没有生成日志文件。", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("打开日志文件夹时发生错误：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 发货运单号列选择按钮
        /// </summary>
        private void btnSelectTrackCol_Click(object sender, EventArgs e)
        {
            SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingTrackColumn, "发货明细运单号列");
        }

        /// <summary>
        /// 发货商品编码列选择按钮
        /// </summary>
        private void btnSelectProductCol_Click(object sender, EventArgs e)
        {
            SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingProductColumn, "发货明细商品编码列");
        }

        /// <summary>
        /// 发货商品名称列选择按钮
        /// </summary>
        private void btnSelectNameCol_Click(object sender, EventArgs e)
        {
            SelectColumnForWorkbook(cmbShippingWorkbook, txtShippingNameColumn, "发货明细商品名称列");
        }

        /// <summary>
        /// 账单运单号列选择按钮
        /// </summary>
        private void btnSelectBillTrackCol_Click(object sender, EventArgs e)
        {
            SelectColumnForWorkbook(cmbBillWorkbook, txtBillTrackColumn, "账单明细运单号列");
        }

        /// <summary>
        /// 账单商品编码列选择按钮
        /// </summary>
        private void btnSelectBillProductCol_Click(object sender, EventArgs e)
        {
            SelectColumnForWorkbook(cmbBillWorkbook, txtBillProductColumn, "账单明细商品编码列");
        }

        /// <summary>
        /// 账单商品名称列选择按钮
        /// </summary>
        private void btnSelectBillNameCol_Click(object sender, EventArgs e)
        {
            SelectColumnForWorkbook(cmbBillWorkbook, txtBillNameColumn, "账单明细商品名称列");
        }

        /// <summary>
        /// 为指定工作簿选择列
        /// </summary>
        private void SelectColumnForWorkbook(ComboBox workbookCombo, TextBox targetTextBox, string title)
        {
            try
            {
                if (workbookCombo.SelectedIndex < 0)
                {
                    MessageBox.Show("请先选择工作簿！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                WorkbookInfo selectedWorkbook = workbooks[workbookCombo.SelectedIndex];
                
                // 暂时隐藏当前窗体
                this.Visible = false;
                
                // 激活选定的工作簿
                selectedWorkbook.Workbook.Activate();
                
                // 显示提示消息并获取用户选择
                MessageBox.Show(string.Format("请在工作簿 [{0}] 中选择 {1}，然后点击确定", 
                    selectedWorkbook.Name, title), "选择列", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 获取用户当前选择的区域
                Excel.Range selection = selectedWorkbook.Workbook.Application.Selection;
                if (selection != null)
                {
                    // 获取选择区域的列字母
                    string columnLetter = ExcelHelper.GetColumnLetter(selection.Column);
                    targetTextBox.Text = columnLetter;
                }
                
                // 恢复窗体显示并聚焦
                this.Visible = true;
                this.WindowState = FormWindowState.Normal;
                this.TopMost = true;
                this.Activate();
                this.Focus();
                this.BringToFront();
                this.TopMost = false;
            }
            catch (Exception ex)
            {
                // 确保窗体可见
                this.Visible = true;
                MessageBox.Show(string.Format("选择列时发生错误：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 设置按钮点击事件
        /// </summary>
        private void btnSettings_Click(object sender, EventArgs e)
        {
            try
            {
                SettingsForm settingsForm = new SettingsForm();
                if (settingsForm.ShowDialog() == DialogResult.OK)
                {
                    // 重新应用设置
                    ApplySettings();
                    SetDefaultValues();
                    
                    MessageBox.Show("设置已应用！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("打开设置窗口失败：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 写入日志的简化方法
        /// </summary>
        private void WriteLog(string message, LogLevel level)
        {
            try
            {
                // 使用MatchService的日志功能
                System.Diagnostics.Debug.WriteLine(string.Format("[{0}] {1}", level, message));
            }
            catch
            {
                // 日志写入失败时不抛出异常
            }
        }

        /// <summary>
        /// 窗体关闭事件
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (isProcessing)
            {
                DialogResult result = MessageBox.Show("正在处理中，确定要退出吗？", "确认", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
                
                backgroundWorker.CancelAsync();
            }
            
            base.OnFormClosing(e);
        }
    }

    /// <summary>
    /// 匹配配置类
    /// </summary>
    public class MatchConfig
    {
        public string ShippingSheetName { get; set; }
        public string BillSheetName { get; set; }
        public string ShippingTrackColumn { get; set; }
        public string ShippingProductColumn { get; set; }
        public string ShippingNameColumn { get; set; }
        public string BillTrackColumn { get; set; }
        public string BillProductColumn { get; set; }
        public string BillNameColumn { get; set; }
    }

    /// <summary>
    /// 多工作簿匹配配置类
    /// </summary>
    public class MultiWorkbookMatchConfig : MatchConfig
    {
        public Excel.Workbook ShippingWorkbook { get; set; }
        public Excel.Workbook BillWorkbook { get; set; }
    }

    /// <summary>
    /// 匹配结果类
    /// </summary>
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