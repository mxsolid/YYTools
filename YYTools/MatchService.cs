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

                progressCallback?.Invoke(10, "正在构建发货明细索引...");
                Dictionary<string, List<ShippingItem>> shippingIndex = BuildShippingIndexFast(shippingSheet, config, progressCallback);
                if (CancellationCheck?.Invoke() == true) { result.ErrorMessage = "任务被用户取消"; return result; }

                progressCallback?.Invoke(50, "正在处理账单明细...");
                ProcessBillDetailsFast(billSheet, config, shippingIndex, result, progressCallback);
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

                for (int batchStart = 0; batchStart < dataRows; batchStart += batchSize)
                {
                    int batchEnd = Math.Min(batchStart + batchSize, dataRows);
                    var batchProductData = new List<string>();
                    var batchNameData = new List<string>();

                    for (int i = batchStart; i < batchEnd; i++)
                    {
                        string billTrackNumber = trackData[i];
                        if (!string.IsNullOrWhiteSpace(billTrackNumber))
                        {
                            string normalizedTrack = NormalizeTrackNumber(billTrackNumber);
                            if (shippingIndex.ContainsKey(normalizedTrack))
                            {
                                matchedCount++;
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

                    WriteBatchData(billSheet, config.BillProductColumn, batchStart + 2, batchProductData);
                    WriteBatchData(billSheet, config.BillNameColumn, batchStart + 2, batchNameData);

                    updatedCells += batchProductData.Count(s => !string.IsNullOrEmpty(s)) + batchNameData.Count(s => !string.IsNullOrEmpty(s));

                    if (CancellationCheck?.Invoke() == true) return;

                    int progress = 50 + (int)(45.0 * batchEnd / dataRows);
                    progressCallback?.Invoke(progress, $"处理进度: {batchEnd}/{dataRows} 行");
                }

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