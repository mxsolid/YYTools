// --- 文件 7: MatchService.cs ---
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
            
            try
            {
                WriteLog("开始执行运单匹配 - 极速模式 v3.0", LogLevel.Info);
                progressCallback?.Invoke(1, "正在优化Excel性能...");
                
                // 性能优化设置
                excelApp.ScreenUpdating = false;
                excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                excelApp.EnableEvents = false;
                excelApp.DisplayStatusBar = false;

                progressCallback?.Invoke(5, "正在获取工作表...");
                Excel.Worksheet shippingSheet = GetWorksheet(config.ShippingWorkbook, config.ShippingSheetName);
                Excel.Worksheet billSheet = GetWorksheet(config.BillWorkbook, config.BillSheetName);

                if (shippingSheet == null || billSheet == null)
                {
                    result.ErrorMessage = $"无法找到指定的工作表: '{config.ShippingSheetName}' 或 '{config.BillSheetName}'";
                    WriteLog($"工作表获取失败: {result.ErrorMessage}", LogLevel.Error);
                    return result;
                }

                // 检查数据量，给出性能警告
                var shippingRange = shippingSheet.UsedRange;
                var billRange = billSheet.UsedRange;
                int shippingRows = shippingRange.Rows.Count;
                int billRows = billRange.Rows.Count;
                
                WriteLog($"数据量统计 - 发货表: {shippingRows:N0} 行, 账单表: {billRows:N0} 行", LogLevel.Info);
                
                if (shippingRows > 100000 || billRows > 100000)
                {
                    WriteLog("检测到大数据量，启用高性能模式", LogLevel.Warning);
                    progressCallback?.Invoke(8, "检测到大数据量，正在启用高性能模式...");
                }

                progressCallback?.Invoke(10, "正在构建发货明细索引...");
                Dictionary<string, List<ShippingItem>> shippingIndex = BuildShippingIndexFast(shippingSheet, config, progressCallback);
                if (CancellationCheck?.Invoke() == true) { result.ErrorMessage = "任务被用户取消"; return result; }

                progressCallback?.Invoke(50, "正在处理账单明细...");
                ProcessBillDetailsFast(billSheet, config, shippingIndex, result, progressCallback);
                if (CancellationCheck?.Invoke() == true) { result.ErrorMessage = "任务被用户取消"; return result; }

                stopwatch.Stop();
                result.Success = true;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                
                WriteLog($"匹配完成 - 处理行数: {result.ProcessedRows:N0}, 匹配成功: {result.MatchedCount:N0}, 耗时: {result.ElapsedSeconds:F2}秒", LogLevel.Info);
                progressCallback?.Invoke(100, "匹配完成！");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                
                WriteLog($"匹配过程中发生严重错误: {ex.ToString()}", LogLevel.Error);
                
                // 用户友好的错误提示
                if (ex.Message.Contains("内存") || ex.Message.Contains("内存不足"))
                {
                    result.ErrorMessage = "处理数据量过大，内存不足。建议分批处理或关闭其他程序释放内存。";
                }
                else if (ex.Message.Contains("COM") || ex.Message.Contains("Excel"))
                {
                    result.ErrorMessage = "Excel连接异常，请检查Excel是否正常运行，或尝试重新打开文件。";
                }
                else
                {
                    result.ErrorMessage = $"处理过程中发生错误：{ex.Message}\n\n请查看日志获取详细信息。";
                }
            }
            finally
            {
                try
                {
                    // 恢复Excel设置
                    excelApp.ScreenUpdating = originalScreenUpdating;
                    excelApp.Calculation = originalCalculation;
                    excelApp.EnableEvents = originalEnableEvents;
                    excelApp.DisplayStatusBar = true;
                }
                catch (Exception ex)
                {
                    WriteLog($"恢复Excel设置失败: {ex.Message}", LogLevel.Warning);
                }
            }
            return result;
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
            catch { return null; }
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

                WriteLog($"开始构建发货索引 - 总行数: {totalRows:N0}", LogLevel.Info);

                // 大数据量优化：分批处理
                const int batchSize = 10000;
                int processedRows = 0;
                
                for (int startRow = 2; startRow <= totalRows; startRow += batchSize)
                {
                    if (CancellationCheck?.Invoke() == true) return index;
                    
                    int endRow = Math.Min(startRow + batchSize - 1, totalRows);
                    int currentBatchSize = endRow - startRow + 1;
                    
                    try
                    {
                        // 分批获取数据，减少内存占用
                        Excel.Range trackRange = shippingSheet.Range[$"{ExcelHelper.GetColumnLetter(trackCol)}{startRow}", $"{ExcelHelper.GetColumnLetter(trackCol)}{endRow}"];
                        Excel.Range productRange = shippingSheet.Range[$"{ExcelHelper.GetColumnLetter(productCol)}{startRow}", $"{ExcelHelper.GetColumnLetter(productCol)}{endRow}"];
                        Excel.Range nameRange = shippingSheet.Range[$"{ExcelHelper.GetColumnLetter(nameCol)}{startRow}", $"{ExcelHelper.GetColumnLetter(nameCol)}{endRow}"];

                        object[,] trackData = trackRange.Value2 as object[,];
                        object[,] productData = productRange.Value2 as object[,];
                        object[,] nameData = nameRange.Value2 as object[,];

                        // 处理当前批次
                        for (int i = 1; i <= currentBatchSize; i++)
                        {
                            string trackNumber = GetArrayValue(trackData, i, 1);
                            if (!string.IsNullOrWhiteSpace(trackNumber))
                            {
                                string normalizedTrack = NormalizeTrackNumber(trackNumber);
                                if (!index.ContainsKey(normalizedTrack))
                                {
                                    index[normalizedTrack] = new List<ShippingItem>();
                                }
                                
                                // 限制每个运单号的最大商品数量，防止内存溢出
                                if (index[normalizedTrack].Count < 100)
                                {
                                    index[normalizedTrack].Add(new ShippingItem
                                    {
                                        ProductCode = GetArrayValue(productData, i, 1),
                                        ProductName = GetArrayValue(nameData, i, 1)
                                    });
                                }
                            }
                        }
                        
                        processedRows += currentBatchSize;
                        
                        // 进度报告
                        int progress = 10 + (int)(40.0 * processedRows / (totalRows - 1));
                        progressCallback?.Invoke(progress, $"构建索引: {processedRows:N0}/{totalRows - 1:N0} 行");
                        
                        // 强制垃圾回收，释放内存
                        if (processedRows % 50000 == 0)
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            WriteLog($"内存清理完成 - 已处理 {processedRows:N0} 行", LogLevel.Info);
                        }
                        
                        // 释放Excel对象引用
                        if (trackRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(trackRange);
                        if (productRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(productRange);
                        if (nameRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(nameRange);
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"处理批次 {startRow}-{endRow} 时发生错误: {ex.Message}", LogLevel.Warning);
                        // 继续处理下一批次
                        continue;
                    }
                }
                
                WriteLog($"发货索引构建完成 - 运单数量: {index.Count:N0}", LogLevel.Info);
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

                Excel.Range trackRange = billSheet.Range[$"{ExcelHelper.GetColumnLetter(trackCol)}2", $"{ExcelHelper.GetColumnLetter(trackCol)}{totalRows}"];
                object[,] trackData = trackRange.Value2 as object[,];

                int dataRows = totalRows - 1;
                object[,] productData = new object[dataRows, 1];
                object[,] nameData = new object[dataRows, 1];
                
                int matchedCount = 0;
                AppSettings settings = AppSettings.Instance;

                for (int i = 1; i <= dataRows; i++)
                {
                    if (i % 100 == 0 && CancellationCheck?.Invoke() == true) return;

                    string billTrackNumber = GetArrayValue(trackData, i, 1);
                    if (!string.IsNullOrWhiteSpace(billTrackNumber))
                    {
                        string normalizedTrack = NormalizeTrackNumber(billTrackNumber);
                        if (shippingIndex.ContainsKey(normalizedTrack))
                        {
                            matchedCount++;
                            List<ShippingItem> matchedItems = shippingIndex[normalizedTrack];
                            
                            var productCodes = matchedItems.Select(item => item.ProductCode.Trim()).Where(c => !string.IsNullOrEmpty(c));
                            var productNames = matchedItems.Select(item => item.ProductName.Trim()).Where(n => !string.IsNullOrEmpty(n));
                            
                            if (settings.RemoveDuplicateItems)
                            {
                                productCodes = productCodes.Distinct();
                                productNames = productNames.Distinct();
                            }
                            
                            if (productCodes.Any()) productData[i - 1, 0] = string.Join(settings.ConcatenationDelimiter, productCodes.ToArray());
                            if (productNames.Any()) nameData[i - 1, 0] = string.Join(settings.ConcatenationDelimiter, productNames.ToArray());
                        }
                    }
                    if (i % 500 == 0 || i == dataRows)
                    {
                        int progress = 50 + (int)(40.0 * i / dataRows);
                        progressCallback?.Invoke(progress, $"匹配进度: {i}/{dataRows} 行");
                    }
                }

                progressCallback?.Invoke(90, "正在高速写入结果...");
                int updatedCells = 0;
                updatedCells += BatchWriteColumn(billSheet, productCol, productData, totalRows, "商品编码");
                updatedCells += BatchWriteColumn(billSheet, nameCol, nameData, totalRows, "商品名称");
                
                result.ProcessedRows = dataRows;
                result.MatchedCount = matchedCount;
                result.UpdatedCells = updatedCells;
            }
            catch (Exception ex) { WriteLog("处理账单失败: " + ex.ToString(), LogLevel.Error); throw; }
        }
        
        private int BatchWriteColumn(Excel.Worksheet worksheet, int column, object[,] data, int totalRows, string columnName)
        {
            try
            {
                if (data == null) return 0;
                string columnLetter = ExcelHelper.GetColumnLetter(column);
                Excel.Range targetRange = worksheet.Range[$"{columnLetter}2", $"{columnLetter}{totalRows}"];
                targetRange.NumberFormat = "@";
                targetRange.Value2 = data;
                return data.Cast<object>().Count(v => v != null);
            }
            catch (Exception ex) { WriteLog($"{columnName}列写入失败: {ex.Message}", LogLevel.Error); return 0; }
        }

        private string GetArrayValue(object[,] array, int row, int col)
        {
            try
            {
                if (array == null || row > array.GetLength(0) || col > array.GetLength(1) || row < 1 || col < 1) return "";
                object value = array[row, col];
                return value?.ToString().Trim() ?? "";
            }
            catch { return ""; }
        }
        
        private string NormalizeTrackNumber(string trackNumber)
        {
            if (string.IsNullOrWhiteSpace(trackNumber)) return "";
            string normalized = trackNumber.Trim();
            if (normalized.Contains("E+") || normalized.Contains("e+"))
            {
                if (decimal.TryParse(normalized, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal decValue))
                {
                    return decValue.ToString();
                }
            }
            return normalized;
        }

        public static void WriteLog(string message, LogLevel level)
        {
            try
            {
                string logDir = AppSettings.Instance.LogDirectory;
                if (string.IsNullOrEmpty(logDir)) logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");

                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, $"YYTools_{DateTime.Now:yyyyMMdd}.log");
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level.ToString().ToUpper()}] {message}";
                File.AppendAllText(logFile, logEntry + Environment.NewLine, System.Text.Encoding.UTF8);
            }
            catch { }
        }

        public static string GetLogFolderPath() => AppSettings.Instance.LogDirectory;
        
        public static void CleanupOldLogs()
        {
            try
            {
                string logDir = AppSettings.Instance.LogDirectory;
                if (!Directory.Exists(logDir)) return;
                var cutoffDate = DateTime.Now.AddDays(-7);
                foreach (var file in Directory.GetFiles(logDir, "YYTools_*.log"))
                {
                    if (new FileInfo(file).CreationTime < cutoffDate)
                    {
                        File.Delete(file);
                    }
                }
            }
            catch { }
        }
    }
}