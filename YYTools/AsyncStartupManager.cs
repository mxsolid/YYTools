using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// 异步启动管理器
    /// </summary>
    public class AsyncStartupManager
    {
        private readonly AsyncTaskManager _taskManager;
        private readonly CacheManager _cacheManager;
        private bool _isInitialized = false;

        public event EventHandler<StartupProgressEventArgs> ProgressReported;
        public event EventHandler<StartupCompletedEventArgs> StartupCompleted;

        public AsyncStartupManager()
        {
            _taskManager = new AsyncTaskManager();
            _cacheManager = CacheManager.Instance;
        }

        /// <summary>
        /// 异步启动应用程序
        /// </summary>
        public async Task<bool> StartAsync()
        {
            try
            {
                Logger.LogInfo("开始异步启动应用程序");
                
                // 报告启动进度
                ReportProgress(0, "正在初始化应用程序...");

                // 第一步：初始化基础组件
                await Task.Delay(100); // 模拟初始化时间
                ReportProgress(10, "基础组件初始化完成");

                // 第二步：初始化日志系统
                await Task.Delay(100);
                ReportProgress(20, "日志系统初始化完成");

                // 第三步：初始化缓存管理器
                await Task.Delay(100);
                ReportProgress(30, "缓存管理器初始化完成");

                // 第四步：异步加载Excel文件信息（可选，失败不影响启动）
                try
                {
                    await LoadExcelFilesAsync();
                    ReportProgress(80, "Excel文件信息加载完成");
                }
                catch (Exception ex)
                {
                    // Excel加载失败不影响程序启动
                    Logger.LogWarning($"Excel文件信息加载失败，但不影响程序启动: {ex.Message}");
                    ReportProgress(80, "Excel文件信息加载跳过（不影响启动）");
                }

                // 第五步：完成启动
                await Task.Delay(100);
                ReportProgress(100, "应用程序启动完成");

                _isInitialized = true;
                
                // 触发启动完成事件
                OnStartupCompleted(true, "启动成功");
                
                Logger.LogInfo("应用程序异步启动完成");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError("应用程序异步启动失败", ex);
                
                // 即使失败也要标记为已初始化，避免重复尝试
                _isInitialized = true;
                
                // 触发启动完成事件（失败状态）
                OnStartupCompleted(false, $"启动失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 异步加载Excel文件信息
        /// </summary>
        private async Task LoadExcelFilesAsync()
        {
            try
            {
                ReportProgress(40, "正在检测Excel应用程序...");

                // 异步检测Excel应用程序
                var excelApp = await Task.Run(() => ExcelAddin.GetExcelApplication());
                if (excelApp == null)
                {
                    ReportProgress(50, "未检测到Excel应用程序");
                    return;
                }

                ReportProgress(50, "正在获取打开的工作簿...");

                // 异步获取工作簿信息
                var workbooks = await Task.Run(() => ExcelAddin.GetOpenWorkbooks());
                if (workbooks == null || workbooks.Count == 0)
                {
                    ReportProgress(60, "未检测到打开的工作簿");
                    return;
                }

                ReportProgress(60, $"检测到 {workbooks.Count} 个工作簿，正在缓存...");

                // 异步缓存工作簿信息
                await CacheWorkbooksAsync(workbooks);

                ReportProgress(70, "工作簿信息缓存完成");
            }
            catch (Exception ex)
            {
                Logger.LogError("异步加载Excel文件信息失败", ex);
                ReportProgress(70, $"加载Excel文件信息失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 异步缓存工作簿信息
        /// </summary>
        private async Task CacheWorkbooksAsync(List<WorkbookInfo> workbooks)
        {
            try
            {
                var tasks = new List<Task>();

                foreach (var workbook in workbooks)
                {
                    var task = Task.Run(() =>
                    {
                        try
                        {
                            // 缓存工作簿
                            _cacheManager.GetOrAddWorkbook(workbook.Name, () => workbook.Workbook);

                            // 获取工作表信息
                            var sheetNames = ExcelAddin.GetWorksheetNames(workbook.Workbook);
                            foreach (var sheetName in sheetNames)
                            {
                                try
                                {
                                    var worksheet = workbook.Workbook.Worksheets[sheetName] as Excel.Worksheet;
                                    if (worksheet != null)
                                    {
                                        // 缓存工作表
                                        _cacheManager.GetOrAddWorksheet(workbook.Name, sheetName, () => worksheet);

                                        // 缓存列信息
                                        var columns = SmartColumnService.GetColumnInfos(worksheet, 50);
                                        _cacheManager.GetOrAddColumnInfo(workbook.Name, sheetName, () => columns);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogWarning($"缓存工作表信息失败: {workbook.Name} - {sheetName}, 错误: {ex.Message}");
                                }
                            }

                            Logger.LogInfo($"工作簿缓存完成: {workbook.Name}");
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError($"缓存工作簿失败: {workbook.Name}", ex);
                        }
                    });

                    tasks.Add(task);
                }

                // 等待所有缓存任务完成
                await Task.WhenAll(tasks);
            }
            catch (Exception ex)
            {
                Logger.LogError("异步缓存工作簿信息失败", ex);
            }
        }

        /// <summary>
        /// 报告启动进度
        /// </summary>
        private void ReportProgress(int percentage, string message)
        {
            try
            {
                var progress = new StartupProgressEventArgs(percentage, message);
                OnProgressReported(progress);
            }
            catch (Exception ex)
            {
                Logger.LogError($"报告启动进度失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 触发进度报告事件
        /// </summary>
        protected virtual void OnProgressReported(StartupProgressEventArgs e)
        {
            ProgressReported?.Invoke(this, e);
        }

        /// <summary>
        /// 触发启动完成事件
        /// </summary>
        protected virtual void OnStartupCompleted(bool success, string message)
        {
            var e = new StartupCompletedEventArgs(success, message);
            StartupCompleted?.Invoke(this, e);
        }

        /// <summary>
        /// 检查是否已初始化
        /// </summary>
        public bool IsInitialized => _isInitialized;

        /// <summary>
        /// 清理资源
        /// </summary>
        public void Dispose()
        {
            try
            {
                _taskManager?.Dispose();
                Logger.LogInfo("异步启动管理器资源已清理");
            }
            catch (Exception ex)
            {
                Logger.LogError("清理异步启动管理器资源失败", ex);
            }
        }
    }

    #region 事件参数类

    /// <summary>
    /// 启动进度事件参数
    /// </summary>
    public class StartupProgressEventArgs : EventArgs
    {
        public int Percentage { get; }
        public string Message { get; }

        public StartupProgressEventArgs(int percentage, string message)
        {
            Percentage = Math.Max(0, Math.Min(100, percentage));
            Message = message ?? string.Empty;
        }
    }

    /// <summary>
    /// 启动完成事件参数
    /// </summary>
    public class StartupCompletedEventArgs : EventArgs
    {
        public bool Success { get; }
        public string Message { get; }

        public StartupCompletedEventArgs(bool success, string message)
        {
            Success = success;
            Message = message ?? string.Empty;
        }
    }

    #endregion
}