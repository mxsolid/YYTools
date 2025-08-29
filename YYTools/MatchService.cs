using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices; // Marshal.ReleaseComObject 需要此引用
using System.Text; // 日志文件编码需要
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    // =================================================================================
    // 核心匹配服务类 (MatchService)
    // =================================================================================

    /// <summary>
    /// 提供Excel数据匹配的核心服务。
    /// 此版本针对大规模数据处理进行了深度优化，重点关注性能、内存使用和稳定性。
    /// </summary>
    public class MatchService
    {
        private static readonly string LogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");
        public Func<bool> CancellationCheck { get; set; }

        public delegate void ProgressReportDelegate(int progress, string message);

        /// <summary>
        /// 执行匹配任务的入口方法。
        /// </summary>
        public MatchResult ExecuteMatch(MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback = null)
        {
            return ExecuteMatchUltraFast(config, progressCallback);
        }

        /// <summary>
        /// 终极优化版的匹配执行逻辑。
        /// </summary>
        private MatchResult ExecuteMatchUltraFast(MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback = null)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            var result = new MatchResult();
            Excel.Application excelApp = config.ShippingWorkbook.Application;

            // 1. 保存原始Excel状态，以便在任务结束时恢复
            bool originalScreenUpdating = excelApp.ScreenUpdating;
            Excel.XlCalculation originalCalculation = excelApp.Calculation;
            bool originalEnableEvents = excelApp.EnableEvents;
            bool originalDisplayStatusBar = excelApp.DisplayStatusBar;
            bool originalDisplayAlerts = excelApp.DisplayAlerts;

            // 声明将在 try 块中使用的COM对象变量，以便在 finally 中可以访问并释放它们
            Excel.Worksheet shippingSheet = null;
            Excel.Worksheet billSheet = null;

            try
            {
                WriteLog("================== 新的匹配任务启动 ==================", LogLevel.Info);
                WriteLog("任务模式: 终极性能优化模式", LogLevel.Info);
                progressCallback?.Invoke(1, "正在优化Excel性能，请稍候...");

                // 2. 极致性能优化：关闭所有不必要的Excel功能
                ExcelHelper.OptimizeExcelPerformance(excelApp);

                progressCallback?.Invoke(5, "正在获取工作表对象...");
                shippingSheet = GetWorksheet(config.ShippingWorkbook, config.ShippingSheetName);
                billSheet = GetWorksheet(config.BillWorkbook, config.BillSheetName);

                if (shippingSheet == null || billSheet == null)
                {
                    result.ErrorMessage = $"致命错误：无法找到指定的工作表 '{config.ShippingSheetName}' 或 '{config.BillSheetName}'。请检查工作表名称是否正确。";
                    WriteLog(result.ErrorMessage, LogLevel.Error);
                    return result;
                }

                WriteLog($"成功获取工作表: '{config.ShippingSheetName}' 和 '{config.BillSheetName}'", LogLevel.Info);


                // 3. 构建发货明细索引 (已进行内存和性能优化)
                progressCallback?.Invoke(10, "正在构建发货明细索引...");
                Dictionary<string, List<ShippingItem>> shippingIndex = BuildShippingIndexOptimized(shippingSheet, config, progressCallback);
                if (CancellationCheck?.Invoke() == true)
                {
                    result.ErrorMessage = "任务被用户取消";
                    return result;
                }

                WriteLog($"发货明细索引构建完成，共计 {shippingIndex.Count:N0} 个唯一的运单号。", LogLevel.Info);


                // 4. 处理账单明细，进行匹配和数据回写 (已进行内存和性能优化)
                progressCallback?.Invoke(50, "正在处理账单明细...");
                ProcessBillDetailsOptimized(billSheet, config, shippingIndex, result, progressCallback);
                if (CancellationCheck?.Invoke() == true)
                {
                    result.ErrorMessage = "任务被用户取消";
                    return result;
                }

                WriteLog("账单明细处理与数据回写完成。", LogLevel.Info);


                stopwatch.Stop();
                result.Success = true;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                progressCallback?.Invoke(100, "任务成功完成！");
                WriteLog($"任务成功完成。处理行数: {result.ProcessedRows:N0}，匹配运单: {result.MatchedCount:N0}，总耗时: {result.ElapsedSeconds:F2} 秒", LogLevel.Info);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.ErrorMessage = $"任务执行过程中发生未处理的异常: {ex.Message}";
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                // 记录详细的异常信息，包括堆栈跟踪，便于排查问题
                WriteLog($"任务执行过程中发生严重错误: {ex.ToString()}", LogLevel.Error);
            }
            finally
            {
                // 5. 关键步骤：无论成功或失败，都必须恢复Excel状态并释放COM对象
                WriteLog("开始恢复Excel原始设置并释放资源。", LogLevel.Info);
                ExcelHelper.RestoreExcelPerformance(excelApp, originalScreenUpdating, originalCalculation, originalEnableEvents, originalDisplayStatusBar, originalDisplayAlerts);

                // 释放COM对象，防止Excel进程残留
                if (shippingSheet != null) Marshal.ReleaseComObject(shippingSheet);
                if (billSheet != null) Marshal.ReleaseComObject(billSheet);

                // 强制垃圾回收，帮助.NET运行时更快地清理COM包装器
                GC.Collect();
                GC.WaitForPendingFinalizers();
                WriteLog("资源释放完成。================== 本次任务结束 ==================", LogLevel.Info);
            }

            return result;
        }

        /// <summary>
        /// 优化的索引构建方法，采用分列读取策略以节省内存。
        /// </summary>
        private Dictionary<string, List<ShippingItem>> BuildShippingIndexOptimized(Excel.Worksheet shippingSheet, MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback)
        {
            var index = new Dictionary<string, List<ShippingItem>>();
            Excel.Range usedRange = null;
            Excel.Range trackColRange = null;
            Excel.Range productColRange = null;
            Excel.Range nameColRange = null;

            try
            {
                usedRange = shippingSheet.UsedRange;
                int totalRows = usedRange.Rows.Count;
                if (totalRows < 2)
                {
                    WriteLog("发货明细表为空或只有标题行，跳过索引构建。", LogLevel.Warning);
                    return index;
                }

                WriteLog($"开始构建索引，发货明细表共 {totalRows:N0} 行。", LogLevel.Info);

                // **内存优化核心**: 分别定义每一列的范围，然后一次性读取一整列。
                // 这样做可以避免加载列与列之间的无关数据，极大地降低了内存消耗。
                trackColRange = shippingSheet.Range[$"{config.ShippingTrackColumn}2:{config.ShippingTrackColumn}{totalRows}"];
                productColRange = shippingSheet.Range[$"{config.ShippingProductColumn}2:{config.ShippingProductColumn}{totalRows}"];
                nameColRange = shippingSheet.Range[$"{config.ShippingNameColumn}2:{config.ShippingNameColumn}{totalRows}"];

                progressCallback?.Invoke(15, $"正在读取 {totalRows:N0} 行发货数据至内存...");

                // 一次性将列数据加载到二维数组中
                object[,] trackValues = trackColRange.Value2 as object[,];
                object[,] productValues = productColRange.Value2 as object[,];
                object[,] nameValues = nameColRange.Value2 as object[,];

                if (trackValues == null)
                {
                    WriteLog("未能从发货明细表读取任何运单号数据，索引为空。", LogLevel.Warning);
                    return index;
                }

                int rows = trackValues.GetLength(0);
                progressCallback?.Invoke(25, "读取完成，正在内存中构建索引...");
                WriteLog($"数据读取完毕，共 {rows:N0} 行待处理。开始在内存中构建索引。", LogLevel.Info);

                for (int i = 1; i <= rows; i++)
                {
                    if (i % 5000 == 0) // 每处理5000行更新一次进度，避免过于频繁
                    {
                        if (CancellationCheck?.Invoke() == true) throw new TaskCanceledException("任务被用户取消。");
                        int progress = 25 + (int)(20.0 * i / rows);
                        progressCallback?.Invoke(progress, $"构建索引: {i}/{rows} 行");
                    }

                    string trackNumber = trackValues[i, 1]?.ToString();
                    if (!string.IsNullOrWhiteSpace(trackNumber))
                    {
                        string normalizedTrack = NormalizeTrackNumber(trackNumber);
                        if (!index.ContainsKey(normalizedTrack))
                        {
                            index[normalizedTrack] = new List<ShippingItem>();
                        }

                        index[normalizedTrack].Add(new ShippingItem
                        {
                            // 安全地获取值，即使其他列数据为空或不存在
                            ProductCode = productValues?[i, 1]?.ToString() ?? "",
                            ProductName = nameValues?[i, 1]?.ToString() ?? ""
                        });
                    }
                }

                return index;
            }
            catch (Exception)
            {
                // 如果出现异常，重新抛出，由外层统一处理和记录
                throw;
            }
            finally
            {
                // 确保此方法内创建的COM对象被释放
                if (nameColRange != null) Marshal.ReleaseComObject(nameColRange);
                if (productColRange != null) Marshal.ReleaseComObject(productColRange);
                if (trackColRange != null) Marshal.ReleaseComObject(trackColRange);
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        /// <summary>
        /// 优化的账单处理方法，采用“计算与IO分离”模式。
        /// </summary>
        private void ProcessBillDetailsOptimized(Excel.Worksheet billSheet, MultiWorkbookMatchConfig config, Dictionary<string, List<ShippingItem>> shippingIndex, MatchResult result,
            ProgressReportDelegate progressCallback)
        {
            Excel.Range usedRange = null;
            Excel.Range trackColRange = null;
            Excel.Range productWriteRange = null;
            Excel.Range nameWriteRange = null;

            try
            {
                usedRange = billSheet.UsedRange;
                int totalRows = usedRange.Rows.Count;
                if (totalRows < 2)
                {
                    WriteLog("账单明细表为空或只有标题行，跳过处理。", LogLevel.Warning);
                    result.ProcessedRows = 0;
                    return;
                }

                WriteLog($"开始处理账单，共 {totalRows:N0} 行。", LogLevel.Info);


                // 1. 一次性读取所有需要匹配的运单号到内存
                progressCallback?.Invoke(55, $"正在读取 {totalRows - 1:N0} 行账单运单号至内存...");
                trackColRange = billSheet.Range[$"{config.BillTrackColumn}2:{config.BillTrackColumn}{totalRows}"];
                object[,] trackValues = trackColRange.Value2 as object[,];

                if (trackValues == null)
                {
                    WriteLog("未能从账单明细表读取任何运单号数据。", LogLevel.Warning);
                    result.ProcessedRows = 0;
                    return;
                }

                int dataRows = trackValues.GetLength(0);
                WriteLog($"读取完成，共 {dataRows:N0} 个运单号。开始并行匹配。", LogLevel.Info);

                // 2. 在内存中创建结果数组
                string[] productResults = new string[dataRows];
                string[] nameResults = new string[dataRows];
                int matchedCount = 0;

                AppSettings settings = AppSettings.Instance;
                int maxThreads = Math.Max(1, Math.Min(settings.MaxThreads, Environment.ProcessorCount));
                var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = maxThreads };

                progressCallback?.Invoke(65, "正在并行匹配数据...");
                WriteLog($"启用并行计算，最大线程数: {maxThreads}", LogLevel.Info);


                // 3. 在纯内存中并行计算，不与Excel交互，以获得最大速度和稳定性
                Parallel.For(0, dataRows, parallelOptions, (i, loopState) =>
                {
                    if (CancellationCheck?.Invoke() == true)
                    {
                        loopState.Stop();
                        return;
                    }

                    // 每20000行检查一次取消状态，避免过于频繁的调用
                    if (i % 20000 == 0 && CancellationCheck?.Invoke() == true)
                    {
                        loopState.Stop();
                        return;
                    }

                    string billTrackNumber = trackValues[i + 1, 1]?.ToString(); // 二维数组索引从1开始
                    if (string.IsNullOrWhiteSpace(billTrackNumber))
                    {
                        productResults[i] = "";
                        nameResults[i] = "";
                        return;
                    }

                    string normalizedTrack = NormalizeTrackNumber(billTrackNumber);
                    if (shippingIndex.TryGetValue(normalizedTrack, out List<ShippingItem> matchedItems))
                    {
                        System.Threading.Interlocked.Increment(ref matchedCount); // 线程安全的计数器

                        var productCodes = matchedItems.Select(item => item.ProductCode?.Trim()).Where(c => !string.IsNullOrEmpty(c));
                        var productNames = matchedItems.Select(item => item.ProductName?.Trim()).Where(n => !string.IsNullOrEmpty(n));

                        // 排序逻辑
                        if (config.SortOption == SortOption.Asc)
                        {
                            productCodes = productCodes.OrderBy(x => x, StringComparer.Ordinal);
                            productNames = productNames.OrderBy(x => x, StringComparer.Ordinal);
                        }
                        else if (config.SortOption == SortOption.Desc)
                        {
                            productCodes = productCodes.OrderByDescending(x => x, StringComparer.Ordinal);
                            productNames = productNames.OrderByDescending(x => x, StringComparer.Ordinal);
                        }

                        // 去重逻辑
                        if (settings.RemoveDuplicateItems)
                        {
                            productCodes = productCodes.Distinct();
                            productNames = productNames.Distinct();
                        }

                        productResults[i] = string.Join(settings.ConcatenationDelimiter, productCodes);
                        nameResults[i] = string.Join(settings.ConcatenationDelimiter, productNames);
                    }
                    else
                    {
                        productResults[i] = "";
                        nameResults[i] = "";
                    }
                });

                if (CancellationCheck?.Invoke() == true) throw new TaskCanceledException("任务被用户取消。");

                WriteLog("并行计算完成。", LogLevel.Info);
                progressCallback?.Invoke(90, "计算完成，正在将结果一次性写回Excel...");

                // 4. 所有计算完成后，一次性将结果数组写回Excel
                productWriteRange = billSheet.Range[$"{config.BillProductColumn}2:{config.BillProductColumn}{dataRows + 1}"];
                nameWriteRange = billSheet.Range[$"{config.BillNameColumn}2:{config.BillNameColumn}{dataRows + 1}"];

                WriteColumnData(productWriteRange, productResults);
                WriteColumnData(nameWriteRange, nameResults);

                result.ProcessedRows = dataRows;
                result.MatchedCount = matchedCount;
                result.UpdatedCells = productResults.Count(s => !string.IsNullOrEmpty(s)) + nameResults.Count(s => !string.IsNullOrEmpty(s));
            }
            catch (Exception)
            {
                throw; // 抛给外层处理
            }
            finally
            {
                // 释放COM对象
                if (nameWriteRange != null) Marshal.ReleaseComObject(nameWriteRange);
                if (productWriteRange != null) Marshal.ReleaseComObject(productWriteRange);
                if (trackColRange != null) Marshal.ReleaseComObject(trackColRange);
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        /// <summary>
        /// 将一维字符串数组数据批量写入到指定的Excel Range中。
        /// </summary>
        private void WriteColumnData(Excel.Range destinationRange, string[] data)
        {
            if (data == null || data.Length == 0) return;
            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                object[,] values = new object[data.Length, 1];
                for (int i = 0; i < data.Length; i++)
                {
                    values[i, 0] = data[i];
                }

                destinationRange.Value2 = values;
                sw.Stop();
                WriteLog($"成功将 {data.Length:N0} 行数据批量写入到 {destinationRange.Address}，耗时: {sw.Elapsed.TotalSeconds:F2} 秒", LogLevel.Info);
            }
            catch (Exception ex)
            {
                WriteLog($"严重错误：批量写入数据到 Range ({destinationRange.Address}) 失败: {ex.Message}", LogLevel.Error);
            }
        }

        // --- 以下是辅助方法 ---

        private Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string sheetName)
        {
            try
            {
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == sheetName) return sheet;
                }

                return null;
            }
            catch (Exception ex)
            {
                WriteLog($"获取工作表 '{sheetName}' 失败: {ex.Message}", LogLevel.Error);
                return null;
            }
        }

        private string NormalizeTrackNumber(string trackNumber)
        {
            return trackNumber.Trim().ToUpperInvariant();
        }

        /// <summary>
        /// 写入日志到文件。
        /// </summary>
        public static void WriteLog(string message, LogLevel level)
        {
            try
            {
                if (!Directory.Exists(LogPath)) Directory.CreateDirectory(LogPath);
                string logFile = Path.Combine(LogPath, $"YYTools_{DateTime.Now:yyyy-MM-dd}.log");
                // 使用带时间戳和级别的格式化日志条目
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level.ToString().ToUpper()}] {message}{Environment.NewLine}";
                // 使用UTF-8编码以支持中文，并避免乱码
                File.AppendAllText(logFile, logEntry, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                // 如果连日志都写不了，只能在调试控制台输出了
                Debug.WriteLine($"Failed to write log: {ex.Message}");
            }
        }
        public static string GetLogFolderPath() => LogPath;

        public static void CleanupOldLogs()
        {
            try
            {
                if (!Directory.Exists(LogPath)) return;

                var logFiles = Directory.GetFiles(LogPath, "YYTools_*.log");
                var cutoffDate = DateTime.Now.AddDays(-30);

                foreach (var logFile in logFiles)
                {
                    try
                    {
                        var fileInfo = new FileInfo(logFile);
                        if (fileInfo.CreationTime < cutoffDate)
                        {
                            fileInfo.Delete();
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}