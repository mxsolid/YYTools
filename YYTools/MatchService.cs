using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    public class MatchService
    {
        private static readonly string LogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");
        public Func<bool> CancellationCheck { get; set; }

        public delegate void ProgressReportDelegate(int progress, string message);

        public MatchResult ExecuteMatch(MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback = null)
        {
            return ExecuteMatchUltraFast(config, progressCallback);
        }

        private MatchResult ExecuteMatchUltraFast(MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback = null)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            var result = new MatchResult();
            Excel.Application excelApp = config.ShippingWorkbook.Application;

            bool originalScreenUpdating = excelApp.ScreenUpdating;
            Excel.XlCalculation originalCalculation = excelApp.Calculation;
            bool originalEnableEvents = excelApp.EnableEvents;
            bool originalDisplayStatusBar = excelApp.DisplayStatusBar;
            bool originalDisplayAlerts = excelApp.DisplayAlerts;

            try
            {
                WriteLog("开始执行匹配任务 - 极速模式", LogLevel.Info);
                progressCallback?.Invoke(1, "正在优化Excel性能...");

                ExcelHelper.OptimizeExcelPerformance(excelApp);

                progressCallback?.Invoke(5, "正在获取工作表...");
                Excel.Worksheet shippingSheet = GetWorksheet(config.ShippingWorkbook, config.ShippingSheetName);
                Excel.Worksheet billSheet = GetWorksheet(config.BillWorkbook, config.BillSheetName);

                if (shippingSheet == null || billSheet == null)
                {
                    result.ErrorMessage = $"无法找到指定的工作表: '{config.ShippingSheetName}' 或 '{config.BillSheetName}'";
                    return result;
                }

                CheckWorksheetSize(shippingSheet, "发货明细", progressCallback);
                CheckWorksheetSize(billSheet, "账单明细", progressCallback);

                // 并行处理发货明细和账单明细
                progressCallback?.Invoke(10, "正在并行构建索引和处理数据...");
                
                var shippingIndex = new Dictionary<string, List<ShippingItem>>();
                var processingTasks = new List<System.Threading.Tasks.Task>();
                var semaphore = new System.Threading.SemaphoreSlim(Math.Min(4, Environment.ProcessorCount));
                
                // 任务1：构建发货明细索引
                var shippingTask = System.Threading.Tasks.Task.Run(async () =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        return BuildShippingIndexFast(shippingSheet, config, progressCallback);
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });
                
                // 任务2：预处理账单明细数据
                var billPreprocessTask = System.Threading.Tasks.Task.Run(async () =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        // 预读取账单明细的关键列数据
                        var billData = PreprocessBillData(billSheet, config);
                        return billData;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });
                
                // 等待两个并行任务完成
                System.Threading.Tasks.Task.WhenAll(shippingTask, billPreprocessTask).Wait();
                
                if (CancellationCheck?.Invoke() == true) { result.ErrorMessage = "任务被用户取消"; return result; }
                
                // 获取结果
                shippingIndex = shippingTask.Result;
                var billData = billPreprocessTask.Result;
                
                progressCallback?.Invoke(60, "正在处理账单明细...");
                
                // 使用预处理的数据处理账单明细
                ProcessBillDetailsWithPreprocessedData(billSheet, config, shippingIndex, billData, result, progressCallback);
                
                if (CancellationCheck?.Invoke() == true) { result.ErrorMessage = "任务被用户取消"; return result; }

                stopwatch.Stop();
                result.Success = true;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                progressCallback?.Invoke(100, "任务完成！");

                WriteLog($"任务完成，处理 {result.ProcessedRows:N0} 行，匹配 {result.MatchedCount:N0} 个运单，耗时 {result.ElapsedSeconds:F2} 秒", LogLevel.Info);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                WriteLog($"任务执行过程中发生错误: {ex.ToString()}", LogLevel.Error);
            }
            finally
            {
                ExcelHelper.RestoreExcelPerformance(excelApp, originalScreenUpdating, originalCalculation);
                try
                {
                    excelApp.EnableEvents = originalEnableEvents;
                    excelApp.DisplayStatusBar = originalDisplayStatusBar;
                    excelApp.DisplayAlerts = originalDisplayAlerts;
                }
                catch { }
            }
            return result;
        }
        
        /// <summary>
        /// 预处理账单明细数据，提高后续处理效率
        /// </summary>
        private Dictionary<int, BillRowData> PreprocessBillData(Excel.Worksheet billSheet, MultiWorkbookMatchConfig config)
        {
            var billData = new Dictionary<int, BillRowData>();
            var usedRange = billSheet.UsedRange;
            int maxRows = usedRange.Rows.Count;
            
            // 获取列号
            int trackColNum = ExcelHelper.GetColumnNumber(config.BillTrackColumn);
            int productColNum = ExcelHelper.GetColumnNumber(config.BillProductColumn);
            int nameColNum = ExcelHelper.GetColumnNumber(config.BillNameColumn);
            
            // 分批预处理，避免内存溢出
            int batchSize = Math.Min(1000, maxRows);
            var tasks = new List<System.Threading.Tasks.Task>();
            var semaphore = new System.Threading.SemaphoreSlim(Math.Min(4, Environment.ProcessorCount));
            
            for (int batchStart = 2; batchStart <= maxRows; batchStart += batchSize)
            {
                int batchEnd = Math.Min(batchStart + batchSize - 1, maxRows);
                int startRow = batchStart;
                int endRow = batchEnd;
                
                tasks.Add(System.Threading.Tasks.Task.Run(async () =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        var batchData = new Dictionary<int, BillRowData>();
                        
                        for (int r = startRow; r <= endRow; r++)
                        {
                            try
                            {
                                string trackNumber = ExcelHelper.GetCellValue(billSheet.Cells[r, trackColNum]);
                                if (!string.IsNullOrWhiteSpace(trackNumber))
                                {
                                    batchData[r] = new BillRowData
                                    {
                                        TrackNumber = trackNumber,
                                        ProductColumn = productColNum,
                                        NameColumn = nameColNum,
                                        RowNumber = r
                                    };
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.LogWarning($"预处理账单行 {r} 失败: {ex.Message}");
                            }
                        }
                        
                        // 线程安全地合并结果
                        lock (billData)
                        {
                            foreach (var kvp in batchData)
                            {
                                billData[kvp.Key] = kvp.Value;
                            }
                        }
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }));
            }
            
            // 等待所有预处理任务完成
            System.Threading.Tasks.Task.WhenAll(tasks).Wait();
            
            return billData;
        }
        
        /// <summary>
        /// 使用预处理的数据处理账单明细
        /// </summary>
        private void ProcessBillDetailsWithPreprocessedData(Excel.Worksheet billSheet, MultiWorkbookMatchConfig config, 
            Dictionary<string, List<ShippingItem>> shippingIndex, Dictionary<int, BillRowData> billData, 
            MatchResult result, ProgressReportDelegate progressCallback)
        {
            int processedRows = 0;
            int matchedCount = 0;
            int updatedCells = 0;
            
            // 使用预处理的数据进行处理
            foreach (var kvp in billData)
            {
                try
                {
                    var rowData = kvp.Value;
                    processedRows++;
                    
                    if (shippingIndex.ContainsKey(rowData.TrackNumber))
                    {
                        matchedCount++;
                        var shippingItems = shippingIndex[rowData.TrackNumber];
                        
                        // 构建要写入的数据
                        var productCodes = shippingItems.Select(i => i.ProductCode).Where(pc => !string.IsNullOrWhiteSpace(pc));
                        var productNames = shippingItems.Select(i => i.ProductName).Where(pn => !string.IsNullOrWhiteSpace(pn));
                        
                        // 写入商品编码列
                        if (rowData.ProductColumn > 0 && productCodes.Any())
                        {
                            string productCodeText = string.Join(config.ConcatenationDelimiter, productCodes);
                            if (config.RemoveDuplicateItems)
                            {
                                productCodeText = string.Join(config.ConcatenationDelimiter, productCodes.Distinct());
                            }
                            billSheet.Cells[rowData.RowNumber, rowData.ProductColumn] = productCodeText;
                            updatedCells++;
                        }
                        
                        // 写入商品名称列
                        if (rowData.NameColumn > 0 && productNames.Any())
                        {
                            string productNameText = string.Join(config.ConcatenationDelimiter, productNames);
                            if (config.RemoveDuplicateItems)
                            {
                                productNameText = string.Join(config.ConcatenationDelimiter, productNames.Distinct());
                            }
                            billSheet.Cells[rowData.RowNumber, rowData.NameColumn] = productNameText;
                            updatedCells++;
                        }
                    }
                    
                    // 报告进度
                    if (processedRows % 100 == 0)
                    {
                        int progress = 60 + (processedRows * 40 / billData.Count);
                        progressCallback?.Invoke(progress, $"正在处理账单明细... 已处理 {processedRows:N0} 行");
                    }
                    
                    if (CancellationCheck?.Invoke() == true) return;
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"处理账单行 {kvp.Key} 失败: {ex.Message}");
                }
            }
            
            result.ProcessedRows = processedRows;
            result.MatchedCount = matchedCount;
            result.UpdatedCells = updatedCells;
        }

        private void CheckWorksheetSize(Excel.Worksheet worksheet, string sheetName, ProgressReportDelegate progressCallback)
        {
            try
            {
                var stats = ExcelHelper.GetWorksheetStats(worksheet);
                if (stats.rows > 100000)
                {
                    string warning = $"⚠️ 警告：{sheetName}工作表包含 {stats.rows:N0} 行数据，处理时间可能较长。";
                    progressCallback?.Invoke(0, warning);
                    WriteLog(warning, LogLevel.Warning);
                }

                if (stats.rows > 500000)
                {
                    string criticalWarning = $"🚨 严重警告：{sheetName}工作表数据量过大 ({stats.rows:N0} 行)，建议分批处理或优化数据结构。";
                    progressCallback?.Invoke(0, criticalWarning);
                    WriteLog(criticalWarning, LogLevel.Error);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"检查工作表大小时发生错误: {ex.Message}", LogLevel.Warning);
            }
        }

        private Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string sheetName)
        {
            try
            {
                if (workbook == null) return null;
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == sheetName) return sheet;
                }
                return null;
            }
            catch (Exception ex)
            {
                WriteLog($"获取工作表失败: {ex.Message}", LogLevel.Error);
                return null;
            }
        }

        private Dictionary<string, List<ShippingItem>> BuildShippingIndexFast(Excel.Worksheet shippingSheet, MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback)
        {
            var index = new Dictionary<string, List<ShippingItem>>();
            try
            {
                Excel.Range usedRange = shippingSheet.UsedRange;
                if (usedRange.Rows.Count < 2) return index;

                int totalRows = usedRange.Rows.Count;
                int trackCol = ExcelHelper.GetColumnNumber(config.ShippingTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.ShippingProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.ShippingNameColumn);

                var batchSize = AppSettings.Instance.BatchSize;
                var trackData = new List<string>();
                var productData = new List<string>();
                var nameData = new List<string>();

                for (int startRow = 2; startRow <= totalRows; startRow += batchSize)
                {
                    int endRow = Math.Min(startRow + batchSize - 1, totalRows);

                    trackData.AddRange(ExcelHelper.GetColumnDataBatch(shippingSheet, config.ShippingTrackColumn, startRow, endRow));
                    productData.AddRange(ExcelHelper.GetColumnDataBatch(shippingSheet, config.ShippingProductColumn, startRow, endRow));
                    nameData.AddRange(ExcelHelper.GetColumnDataBatch(shippingSheet, config.ShippingNameColumn, startRow, endRow));

                    if (CancellationCheck?.Invoke() == true) return index;

                    int progress = 10 + (int)(35.0 * (startRow - 2) / (totalRows - 1));
                    progressCallback?.Invoke(progress, $"构建索引: {endRow}/{totalRows} 行");
                }

                for (int i = 0; i < trackData.Count; i++)
                {
                    // 过滤掉空行（若原文件存在大量空行，防止参与后续处理导致卡顿）
                    string trackNumber = trackData[i];
                    bool isRowEmpty = string.IsNullOrWhiteSpace(trackNumber)
                                      && (i >= productData.Count || string.IsNullOrWhiteSpace(productData[i]))
                                      && (i >= nameData.Count || string.IsNullOrWhiteSpace(nameData[i]));
                    if (isRowEmpty) continue;

                    if (!string.IsNullOrWhiteSpace(trackNumber))
                    {
                        string normalizedTrack = NormalizeTrackNumber(trackNumber);
                        if (!index.ContainsKey(normalizedTrack))
                        {
                            index[normalizedTrack] = new List<ShippingItem>();
                        }
                        index[normalizedTrack].Add(new ShippingItem
                        {
                            ProductCode = i < productData.Count ? productData[i] : "",
                            ProductName = i < nameData.Count ? nameData[i] : ""
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("构建索引失败: " + ex.ToString(), LogLevel.Error);
                throw;
            }
            return index;
        }

        private void ProcessBillDetailsFast(Excel.Worksheet billSheet, MultiWorkbookMatchConfig config, Dictionary<string, List<ShippingItem>> shippingIndex, MatchResult result, ProgressReportDelegate progressCallback)
        {
            try
            {
                Excel.Range usedRange = billSheet.UsedRange;
                if (usedRange.Rows.Count < 2) return;

                int totalRows = usedRange.Rows.Count;
                int trackCol = ExcelHelper.GetColumnNumber(config.BillTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.BillProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.BillNameColumn);

                var batchSize = AppSettings.Instance.BatchSize;
                var trackData = new List<string>();

                for (int startRow = 2; startRow <= totalRows; startRow += batchSize)
                {
                    int endRow = Math.Min(startRow + batchSize - 1, totalRows);
                    trackData.AddRange(ExcelHelper.GetColumnDataBatch(billSheet, config.BillTrackColumn, startRow, endRow));

                    if (CancellationCheck?.Invoke() == true) return;
                }

                int dataRows = trackData.Count;
                int matchedCount = 0;
                int updatedCells = 0;
                AppSettings settings = AppSettings.Instance;

                // 使用并行批处理：外层批次循环并发执行，批次内仍为顺序逻辑，保持写回一致性
                int maxThreads = Math.Max(1, Math.Min(AppSettings.Instance.MaxThreads, Environment.ProcessorCount));
                var parallelOptions = new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = maxThreads, CancellationToken = default };

                object aggregateLock = new object();
                int processedRows = 0;

                System.Threading.Tasks.Parallel.For(0, (int)Math.Ceiling(dataRows / (double)batchSize), parallelOptions, (batchIndex) =>
                {
                    if (CancellationCheck?.Invoke() == true) return;

                    int batchStart = batchIndex * batchSize;
                    int batchEnd = Math.Min(batchStart + batchSize, dataRows);
                    var batchProductData = new List<string>(batchEnd - batchStart);
                    var batchNameData = new List<string>(batchEnd - batchStart);

                    int localMatched = 0;

                    for (int i = batchStart; i < batchEnd; i++)
                    {
                        string billTrackNumber = trackData[i];
                        if (!string.IsNullOrWhiteSpace(billTrackNumber))
                        {
                            string normalizedTrack = NormalizeTrackNumber(billTrackNumber);
                            if (shippingIndex.ContainsKey(normalizedTrack))
                            {
                                localMatched++;
                                List<ShippingItem> matchedItems = shippingIndex[normalizedTrack];

                                var productCodes = matchedItems.Select(item => item.ProductCode.Trim()).Where(c => !string.IsNullOrEmpty(c));
                                var productNames = matchedItems.Select(item => item.ProductName.Trim()).Where(n => !string.IsNullOrEmpty(n));

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

                                if (settings.RemoveDuplicateItems)
                                {
                                    productCodes = productCodes.Distinct();
                                    productNames = productNames.Distinct();
                                }

                                batchProductData.Add(productCodes.Any() ? string.Join(settings.ConcatenationDelimiter, productCodes.ToArray()) : "");
                                batchNameData.Add(productNames.Any() ? string.Join(settings.ConcatenationDelimiter, productNames.ToArray()) : "");
                            }
                            else
                            {
                                batchProductData.Add("");
                                batchNameData.Add("");
                            }
                        }
                        else
                        {
                            batchProductData.Add("");
                            batchNameData.Add("");
                        }
                    }

                    // 批量写回到Excel（COM对象非线程安全，序列化写入）
                    lock (aggregateLock)
                    {
                        WriteBatchData(billSheet, config.BillProductColumn, batchStart + 2, batchProductData);
                        WriteBatchData(billSheet, config.BillNameColumn, batchStart + 2, batchNameData);

                        matchedCount += localMatched;
                        updatedCells += batchProductData.Count(s => !string.IsNullOrEmpty(s)) + batchNameData.Count(s => !string.IsNullOrEmpty(s));
                        processedRows = Math.Max(processedRows, batchEnd);

                        int progress = 50 + (int)(45.0 * processedRows / dataRows);
                        progressCallback?.Invoke(progress, $"处理进度: {processedRows}/{dataRows} 行");
                    }
                });

                result.ProcessedRows = dataRows;
                result.MatchedCount = matchedCount;
                result.UpdatedCells = updatedCells;
            }
            catch (Exception ex)
            {
                WriteLog($"处理账单明细失败: {ex.Message}", LogLevel.Error);
                throw;
            }
        }

        private void WriteBatchData(Excel.Worksheet worksheet, string columnLetter, int startRow, List<string> data)
        {
            try
            {
                if (data.Count == 0 || string.IsNullOrWhiteSpace(columnLetter) || !ExcelHelper.IsValidColumnLetter(columnLetter)) return;

                var range = worksheet.Range[$"{columnLetter}{startRow}:{columnLetter}{startRow + data.Count - 1}"];
                if (range == null) return;

                var values = new object[data.Count, 1];
                for (int i = 0; i < data.Count; i++)
                {
                    values[i, 0] = data[i];
                }

                range.Value2 = values;
            }
            catch (Exception ex)
            {
                WriteLog($"批量写入数据失败: {ex.Message}", LogLevel.Error);
            }
        }

        private string NormalizeTrackNumber(string trackNumber)
        {
            if (string.IsNullOrWhiteSpace(trackNumber)) return "";
            return trackNumber.Trim().ToUpperInvariant();
        }

        private string GetArrayValue(object[,] array, int row, int col)
        {
            try
            {
                if (array != null && row <= array.GetLength(0) && col <= array.GetLength(1))
                {
                    return array[row, col]?.ToString().Trim() ?? "";
                }
            }
            catch { }
            return "";
        }

        public static void WriteLog(string message, LogLevel level)
        {
            try
            {
                if (!Directory.Exists(LogPath)) Directory.CreateDirectory(LogPath);
                string logFile = Path.Combine(LogPath, $"YYTools_{DateTime.Now:yyyy-MM-dd}.log");
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                File.AppendAllText(logFile, logEntry + Environment.NewLine, System.Text.Encoding.UTF8);
            }
            catch { }
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