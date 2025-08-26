using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// 运单匹配服务类，负责核心的匹配逻辑
    /// </summary>
    public class MatchService
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "YYTools", "Logs");

        /// <summary>
        /// 进度报告委托
        /// </summary>
        /// <param name="progress">进度百分比(0-100)</param>
        /// <param name="message">状态消息</param>
        public delegate void ProgressReportDelegate(int progress, string message);

        /// <summary>
        /// 执行运单匹配操作 - 智能性能模式
        /// </summary>
        public MatchResult ExecuteMatch(MultiWorkbookMatchConfig config,
            ProgressReportDelegate progressCallback = null)
        {
            // 获取性能设置并进行硬件优化
            AppSettings settings = AppSettings.Instance;
            PerformanceMode optimizedMode = OptimizePerformanceMode(settings.PerformanceMode);
            
            switch (optimizedMode)
            {
                case PerformanceMode.UltraFast:
                    return ExecuteMatchUltraFast(config, progressCallback);
                case PerformanceMode.Balanced:
                    return ExecuteMatchBalancedOptimized(config, progressCallback);
                case PerformanceMode.Compatible:
                    return ExecuteMatchCompatible(config, progressCallback);
                default:
                    return ExecuteMatchUltraFast(config, progressCallback);
            }
        }
        
        /// <summary>
        /// 根据硬件配置优化性能模式
        /// </summary>
        private PerformanceMode OptimizePerformanceMode(PerformanceMode selectedMode)
        {
            try
            {
                // 获取硬件信息
                int coreCount = Environment.ProcessorCount;
                long memoryMB = GC.GetTotalMemory(false) / 1024 / 1024;
                
                WriteLog(string.Format("硬件检测 - CPU核心数: {0}, 可用内存: {1}MB", coreCount, memoryMB), LogLevel.Info);
                
                // 如果选择平衡模式，根据硬件进行智能优化
                if (selectedMode == PerformanceMode.Balanced)
                {
                    if (coreCount >= 8 && memoryMB > 4096) // 高配置机器
                    {
                        WriteLog("检测到高配置硬件，自动升级到极速模式", LogLevel.Info);
                        return PerformanceMode.UltraFast;
                    }
                    else if (coreCount <= 2 || memoryMB < 2048) // 低配置机器
                    {
                        WriteLog("检测到低配置硬件，自动降级到兼容模式", LogLevel.Info);
                        return PerformanceMode.Compatible;
                    }
                }
                
                return selectedMode;
            }
            catch (Exception ex)
            {
                WriteLog("硬件检测失败: " + ex.Message, LogLevel.Warning);
                return selectedMode;
            }
        }

        /// <summary>
        /// 极速模式 - 最高性能
        /// </summary>
        private MatchResult ExecuteMatchUltraFast(MultiWorkbookMatchConfig config,
            ProgressReportDelegate progressCallback = null)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            MatchResult result = new MatchResult();
            
            Excel.Application excelApp = config.ShippingWorkbook.Application;

            // Excel性能优化设置
            bool originalScreenUpdating = true;
            bool originalDisplayAlerts = true;
            Excel.XlCalculation originalCalculation = Excel.XlCalculation.xlCalculationAutomatic;
            bool originalEnableEvents = true;
            
            try
            {
                WriteLog("开始执行运单匹配 - 极速模式", LogLevel.Info);
                WriteLog(string.Format("配置信息 - 发货表:{0}, 账单表:{1}", 
                    config.ShippingSheetName, config.BillSheetName), LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(1, "正在优化Excel性能设置...");

                // 保存原始设置并优化性能
                originalScreenUpdating = excelApp.ScreenUpdating;
                originalDisplayAlerts = excelApp.DisplayAlerts;
                originalCalculation = excelApp.Calculation;
                originalEnableEvents = excelApp.EnableEvents;
                
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;
                excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                excelApp.EnableEvents = false;

                if (progressCallback != null)
                    progressCallback(5, "正在获取工作表...");

                Excel.Worksheet shippingSheet = GetWorksheet(config.ShippingWorkbook, config.ShippingSheetName);
                Excel.Worksheet billSheet = GetWorksheet(config.BillWorkbook, config.BillSheetName);

                if (shippingSheet == null || billSheet == null)
                {
                    result.ErrorMessage = $"无法找到指定的工作表: '{config.ShippingSheetName}' 或 '{config.BillSheetName}'";
                    WriteLog("工作表获取失败: " + result.ErrorMessage, LogLevel.Error);
                    return result;
                }

                if (progressCallback != null)
                    progressCallback(10, "正在构建发货明细索引...");

                // 构建发货明细索引 - 使用批量读取
                Dictionary<string, List<ShippingItem>> shippingIndex = BuildShippingIndexFast(
                    shippingSheet, config, progressCallback);

                WriteLog(string.Format("发货明细索引构建完成，共{0}个运单号", shippingIndex.Count), LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(50, "正在处理账单明细...");

                // 处理账单明细 - 使用批量处理
                ProcessBillDetailsFast(billSheet, config, shippingIndex, result, progressCallback);

                stopwatch.Stop();
                result.Success = true;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;

                WriteLog(string.Format("匹配完成 - 处理{0}行，匹配{1}个，更新{2}个单元格，耗时{3:F2}秒", 
                    result.ProcessedRows, result.MatchedCount, result.UpdatedCells, result.ElapsedSeconds), LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(100, "匹配完成！");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                
                WriteLog("匹配过程中发生错误: " + ex.ToString(), LogLevel.Error);
            }
            finally
            {
                // 恢复Excel设置
                try
                {
                    excelApp.ScreenUpdating = originalScreenUpdating;
                    excelApp.DisplayAlerts = originalDisplayAlerts;
                    excelApp.Calculation = originalCalculation;
                    excelApp.EnableEvents = originalEnableEvents;
                }
                catch (Exception ex)
                {
                    WriteLog("恢复Excel设置失败: " + ex.Message, LogLevel.Warning);
                }
            }

            return result;
        }

        /// <summary>
        /// 平衡模式优化版 - 智能批处理
        /// </summary>
        private MatchResult ExecuteMatchBalancedOptimized(MultiWorkbookMatchConfig config, 
            ProgressReportDelegate progressCallback = null)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            MatchResult result = new MatchResult();
            Excel.Application excelApp = config.ShippingWorkbook.Application;

            // 轻量级Excel优化
            bool originalScreenUpdating = true;
            Excel.XlCalculation originalCalculation = Excel.XlCalculation.xlCalculationAutomatic;
            
            try
            {
                WriteLog("开始执行运单匹配 - 平衡模式优化版", LogLevel.Info);

                // 轻量级性能优化
                originalScreenUpdating = excelApp.ScreenUpdating;
                originalCalculation = excelApp.Calculation;
                excelApp.ScreenUpdating = false;
                excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                if (progressCallback != null)
                    progressCallback(5, "正在获取工作表...");

                // 获取工作表
                Excel.Worksheet shippingSheet = GetWorksheet(config.ShippingWorkbook, config.ShippingSheetName);
                Excel.Worksheet billSheet = GetWorksheet(config.BillWorkbook, config.BillSheetName);

                if (shippingSheet == null || billSheet == null)
                {
                    result.ErrorMessage = "无法找到指定的工作表";
                    return result;
                }

                if (progressCallback != null)
                    progressCallback(15, "正在构建发货明细索引...");

                // 使用中等批量读取 - 根据硬件调整批次大小
                int batchSize = CalculateOptimalBatchSize();
                Dictionary<string, List<ShippingItem>> shippingIndex = BuildShippingIndexBalancedOptimized(
                    shippingSheet, config, batchSize, progressCallback);

                if (progressCallback != null)
                    progressCallback(60, "正在处理账单明细...");

                // 使用优化的平衡处理
                ProcessBillDetailsBalancedOptimized(billSheet, config, shippingIndex, result, batchSize, progressCallback);

                stopwatch.Stop();
                result.Success = true;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;

                WriteLog(string.Format("平衡模式优化版匹配完成 - 耗时{0:F2}秒", result.ElapsedSeconds), LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(100, "匹配完成！");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                WriteLog("平衡模式优化版匹配错误: " + ex.ToString(), LogLevel.Error);
            }
            finally
            {
                // 恢复Excel设置
                try
                {
                    excelApp.ScreenUpdating = originalScreenUpdating;
                    excelApp.Calculation = originalCalculation;
                }
                catch (Exception ex)
                {
                    WriteLog("恢复Excel设置失败: " + ex.Message, LogLevel.Warning);
                }
            }

            return result;
        }

        /// <summary>
        /// 兼容模式 - 最佳兼容性
        /// </summary>
        private MatchResult ExecuteMatchCompatible(MultiWorkbookMatchConfig config,
            ProgressReportDelegate progressCallback = null)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            MatchResult result = new MatchResult();
            
            try
            {
                WriteLog("开始执行运单匹配 - 兼容模式", LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(5, "正在获取工作表...");

                // 获取工作表
                Excel.Worksheet shippingSheet = GetWorksheet(config.ShippingWorkbook, config.ShippingSheetName);
                Excel.Worksheet billSheet = GetWorksheet(config.BillWorkbook, config.BillSheetName);

                if (shippingSheet == null || billSheet == null)
                {
                    result.ErrorMessage = "无法找到指定的工作表";
                    return result;
                }

                if (progressCallback != null)
                    progressCallback(15, "正在构建发货明细索引...");

                // 使用传统逐行读取
                Dictionary<string, List<ShippingItem>> shippingIndex = BuildShippingIndexCompatible(
                    shippingSheet, config, progressCallback);

                if (progressCallback != null)
                    progressCallback(60, "正在处理账单明细...");

                // 使用传统逐行处理
                ProcessBillDetailsCompatible(billSheet, config, shippingIndex, result, progressCallback);

                stopwatch.Stop();
                result.Success = true;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;

                WriteLog(string.Format("兼容模式匹配完成 - 耗时{0:F2}秒", result.ElapsedSeconds), LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(100, "匹配完成！");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.ElapsedSeconds = stopwatch.Elapsed.TotalSeconds;
                WriteLog("兼容模式匹配错误: " + ex.ToString(), LogLevel.Error);
            }

            return result;
        }

        /// <summary>
        /// 获取工作表 (在指定工作簿中查找)
        /// </summary>
        private Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string sheetName)
        {
            try
            {
                if (workbook == null)
                {
                    WriteLog("传递给 GetWorksheet 的工作簿为 null", LogLevel.Error);
                    return null;
                }

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == sheetName)
                        return sheet;
                }
                WriteLog($"在工作簿 '{workbook.Name}' 中未找到名为 '{sheetName}' 的工作表", LogLevel.Warning);
                return null;
            }
            catch (Exception ex)
            {
                WriteLog("获取工作表失败: " + ex.Message, LogLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// 高性能发货明细索引构建 - 批量读取
        /// </summary>
        private Dictionary<string, List<ShippingItem>> BuildShippingIndexFast(
            Excel.Worksheet shippingSheet, MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback)
        {
            Dictionary<string, List<ShippingItem>> index = new Dictionary<string, List<ShippingItem>>();

            try
            {
                Excel.Range usedRange = shippingSheet.UsedRange;
                if (usedRange == null || usedRange.Rows.Count < 2)
                {
                    WriteLog("发货明细表为空或只有标题行", LogLevel.Warning);
                    return index;
                }

                int totalRows = usedRange.Rows.Count;
                WriteLog(string.Format("发货明细表共 {0} 行数据", totalRows), LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(12, "正在批量读取发货数据...");
                
                // --- FIX: Use the new method name from the user-provided ExcelHelper.
                int trackCol = ExcelHelper.GetColumnNumber(config.ShippingTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.ShippingProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.ShippingNameColumn);

                Excel.Range trackRange = shippingSheet.Range[
                    string.Format("{0}2", ExcelHelper.GetColumnLetter(trackCol)),
                    string.Format("{0}{1}", ExcelHelper.GetColumnLetter(trackCol), totalRows)];
                Excel.Range productRange = shippingSheet.Range[
                    string.Format("{0}2", ExcelHelper.GetColumnLetter(productCol)),
                    string.Format("{0}{1}", ExcelHelper.GetColumnLetter(productCol), totalRows)];
                Excel.Range nameRange = shippingSheet.Range[
                    string.Format("{0}2", ExcelHelper.GetColumnLetter(nameCol)),
                    string.Format("{0}{1}", ExcelHelper.GetColumnLetter(nameCol), totalRows)];

                object[,] trackData = trackRange.Value2 as object[,];
                object[,] productData = productRange.Value2 as object[,];
                object[,] nameData = nameRange.Value2 as object[,];

                if (progressCallback != null)
                    progressCallback(25, "正在处理发货数据索引...");

                int validCount = 0;
                int dataRows = totalRows - 1; 

                index = new Dictionary<string, List<ShippingItem>>(dataRows);

                for (int i = 1; i <= dataRows; i++)
                {
                    string trackNumber = GetArrayValue(trackData, i, 1);
                    string productCode = GetArrayValue(productData, i, 1);
                    string productName = GetArrayValue(nameData, i, 1);

                    if (!string.IsNullOrWhiteSpace(trackNumber))
                    {
                        string normalizedTrack = NormalizeTrackNumber(trackNumber);

                        if (!index.ContainsKey(normalizedTrack))
                        {
                            index[normalizedTrack] = new List<ShippingItem>();
                        }

                        index[normalizedTrack].Add(new ShippingItem
                        {
                            ProductCode = productCode ?? "",
                            ProductName = productName ?? ""
                        });

                        validCount++;
                    }

                    if (i % 1000 == 0 || i == dataRows)
                    {
                        int progress = 25 + (int)(20.0 * i / dataRows);
                        if (progressCallback != null)
                            progressCallback(progress, string.Format("构建索引: {0}/{1} 行 ({2} 个有效运单)", 
                                i, dataRows, validCount));
                    }
                }

                WriteLog(string.Format("发货明细索引构建完成 - 有效运单: {0} 个", index.Count), LogLevel.Info);
            }
            catch (Exception ex)
            {
                WriteLog("构建发货明细索引错误: " + ex.ToString(), LogLevel.Error);
                throw new Exception("构建发货明细索引失败: " + ex.Message);
            }

            return index;
        }

        /// <summary>
        /// 终极性能账单明细处理
        /// </summary>
        private void ProcessBillDetailsFast(Excel.Worksheet billSheet, MultiWorkbookMatchConfig config,
            Dictionary<string, List<ShippingItem>> shippingIndex, MatchResult result, 
            ProgressReportDelegate progressCallback)
        {
            try
            {
                Excel.Range usedRange = billSheet.UsedRange;
                if (usedRange == null || usedRange.Rows.Count < 2)
                {
                    WriteLog("账单明细表为空或只有标题行", LogLevel.Warning);
                    return;
                }

                int totalRows = usedRange.Rows.Count;
                WriteLog(string.Format("账单明细表共 {0} 行数据", totalRows), LogLevel.Info);

                if (progressCallback != null)
                    progressCallback(52, "正在批量读取账单数据...");
                
                // --- FIX: Use the new method name from the user-provided ExcelHelper.
                int trackCol = ExcelHelper.GetColumnNumber(config.BillTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.BillProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.BillNameColumn);

                Excel.Range trackRange = billSheet.Range[
                    string.Format("{0}2", ExcelHelper.GetColumnLetter(trackCol)),
                    string.Format("{0}{1}", ExcelHelper.GetColumnLetter(trackCol), totalRows)];
                object[,] trackData = trackRange.Value2 as object[,];

                if (trackData == null)
                {
                    WriteLog("无法读取账单运单号列数据", LogLevel.Error);
                    throw new Exception("无法读取账单运单号列数据");
                }

                if (progressCallback != null)
                    progressCallback(60, "开始匹配处理...");

                int matchedCount = 0;
                int processedRows = 0;
                int dataRows = totalRows - 1; 

                object[,] productData = new object[dataRows, 1];
                object[,] nameData = new object[dataRows, 1];
                bool[] hasProductUpdate = new bool[dataRows];
                bool[] hasNameUpdate = new bool[dataRows];

                for (int i = 1; i <= dataRows; i++)
                {
                    processedRows++;
                    
                    string billTrackNumber = GetArrayValue(trackData, i, 1);

                    if (!string.IsNullOrWhiteSpace(billTrackNumber))
                    {
                        string normalizedTrack = NormalizeTrackNumber(billTrackNumber);

                        if (shippingIndex.ContainsKey(normalizedTrack))
                        {
                            List<ShippingItem> matchedItems = shippingIndex[normalizedTrack];
                            matchedCount++;

                            HashSet<string> productCodes = new HashSet<string>();
                            HashSet<string> productNames = new HashSet<string>();

                            foreach (ShippingItem item in matchedItems)
                            {
                                if (!string.IsNullOrWhiteSpace(item.ProductCode))
                                    productCodes.Add(item.ProductCode.Trim());
                                if (!string.IsNullOrWhiteSpace(item.ProductName))
                                    productNames.Add(item.ProductName.Trim());
                            }

                            int arrayIndex = i - 1; 
                            if (productCodes.Count > 0)
                            {
                                productData[arrayIndex, 0] = string.Join("、", productCodes.ToArray());
                                hasProductUpdate[arrayIndex] = true;
                            }

                            if (productNames.Count > 0)
                            {
                                nameData[arrayIndex, 0] = string.Join("、", productNames.ToArray());
                                hasNameUpdate[arrayIndex] = true;
                            }
                        }
                    }

                    if (processedRows % 500 == 0 || i == dataRows)
                    {
                        int progress = 60 + (int)(25.0 * processedRows / dataRows);
                        if (progressCallback != null)
                            progressCallback(progress, string.Format("匹配进度: {0}/{1} 行 (已匹配 {2} 个)", 
                                processedRows, dataRows, matchedCount));
                    }
                }

                if (progressCallback != null)
                    progressCallback(88, "正在超高速批量写入结果...");

                int updatedCells = 0;
                updatedCells += UltimateBatchWrite(billSheet, productCol, productData, hasProductUpdate, dataRows, totalRows, "商品编码");
                updatedCells += UltimateBatchWrite(billSheet, nameCol, nameData, hasNameUpdate, dataRows, totalRows, "商品名称");

                result.ProcessedRows = processedRows;
                result.MatchedCount = matchedCount;
                result.UpdatedCells = updatedCells;

                WriteLog(string.Format("账单处理完成 - 匹配 {0} 个，更新 {1} 个单元格", 
                    matchedCount, updatedCells), LogLevel.Info);
            }
            catch (Exception ex)
            {
                WriteLog("处理账单明细错误: " + ex.ToString(), LogLevel.Error);
                throw new Exception("处理账单明细失败: " + ex.Message);
            }
        }

        /// <summary>
        /// 平衡模式优化版索引构建
        /// </summary>
        private Dictionary<string, List<ShippingItem>> BuildShippingIndexBalancedOptimized(
            Excel.Worksheet shippingSheet, MultiWorkbookMatchConfig config, int batchSize, ProgressReportDelegate progressCallback)
        {
            Dictionary<string, List<ShippingItem>> index = new Dictionary<string, List<ShippingItem>>();

            try
            {
                Excel.Range usedRange = shippingSheet.UsedRange;
                if (usedRange == null || usedRange.Rows.Count < 2) return index;

                int totalRows = usedRange.Rows.Count;
                WriteLog(string.Format("平衡模式优化版 - 批处理大小: {0}, 总行数: {1}", batchSize, totalRows), LogLevel.Info);

                // --- FIX: Use the new method name from the user-provided ExcelHelper.
                int trackCol = ExcelHelper.GetColumnNumber(config.ShippingTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.ShippingProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.ShippingNameColumn);
                
                string trackColLetter = ExcelHelper.GetColumnLetter(trackCol);
                string productColLetter = ExcelHelper.GetColumnLetter(productCol);
                string nameColLetter = ExcelHelper.GetColumnLetter(nameCol);

                index = new Dictionary<string, List<ShippingItem>>(totalRows);

                for (int startRow = 2; startRow <= totalRows; startRow += batchSize)
                {
                    int endRow = Math.Min(startRow + batchSize - 1, totalRows);
                    
                    Excel.Range trackRange = shippingSheet.Range[$"{trackColLetter}{startRow}", $"{trackColLetter}{endRow}"];
                    Excel.Range productRange = shippingSheet.Range[$"{productColLetter}{startRow}", $"{productColLetter}{endRow}"];
                    Excel.Range nameRange = shippingSheet.Range[$"{nameColLetter}{startRow}", $"{nameColLetter}{endRow}"];

                    object[,] trackData = trackRange.Value2 as object[,];
                    object[,] productData = productRange.Value2 as object[,];
                    object[,] nameData = nameRange.Value2 as object[,];

                    int batchRows = endRow - startRow + 1;
                    for (int i = 1; i <= batchRows; i++)
                    {
                        string trackNumber = GetArrayValue(trackData, i, 1);
                        string productCode = GetArrayValue(productData, i, 1);
                        string productName = GetArrayValue(nameData, i, 1);

                        if (!string.IsNullOrWhiteSpace(trackNumber))
                        {
                            string normalizedTrack = NormalizeTrackNumber(trackNumber);
                            if (!index.ContainsKey(normalizedTrack))
                                index[normalizedTrack] = new List<ShippingItem>();

                            index[normalizedTrack].Add(new ShippingItem
                            {
                                ProductCode = productCode ?? "",
                                ProductName = productName ?? ""
                            });
                        }
                    }

                    if (progressCallback != null)
                    {
                        int progress = 15 + (int)(40.0 * endRow / totalRows);
                        progressCallback(progress, string.Format("智能批处理索引: {0}/{1} 行 (批量: {2})", 
                            endRow, totalRows, batchSize));
                    }
                }

                WriteLog(string.Format("平衡模式优化版索引构建完成 - 有效运单: {0} 个", index.Count), LogLevel.Info);
            }
            catch (Exception ex)
            {
                WriteLog("平衡模式优化版索引构建错误: " + ex.ToString(), LogLevel.Error);
                throw;
            }

            return index;
        }

        /// <summary>
        /// 平衡模式优化版账单处理
        /// </summary>
        private void ProcessBillDetailsBalancedOptimized(Excel.Worksheet billSheet, MultiWorkbookMatchConfig config,
            Dictionary<string, List<ShippingItem>> shippingIndex, MatchResult result, int batchSize,
            ProgressReportDelegate progressCallback)
        {
            try
            {
                Excel.Range usedRange = billSheet.UsedRange;
                if (usedRange == null || usedRange.Rows.Count < 2) return;

                int totalRows = usedRange.Rows.Count;
                int matchedCount = 0;
                int processedRows = 0;

                // --- FIX: Use the new method name from the user-provided ExcelHelper.
                int trackCol = ExcelHelper.GetColumnNumber(config.BillTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.BillProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.BillNameColumn);
                string trackColLetter = ExcelHelper.GetColumnLetter(trackCol);

                int writeBatchSize = batchSize / 2; 
                Dictionary<int, string> productUpdates = new Dictionary<int, string>();
                Dictionary<int, string> nameUpdates = new Dictionary<int, string>();

                for (int startRow = 2; startRow <= totalRows; startRow += batchSize)
                {
                    int endRow = Math.Min(startRow + batchSize - 1, totalRows);
                    
                    Excel.Range trackRange = billSheet.Range[$"{trackColLetter}{startRow}", $"{trackColLetter}{endRow}"];
                    object[,] trackData = trackRange.Value2 as object[,];

                    int batchRows = endRow - startRow + 1;
                    for (int i = 1; i <= batchRows; i++)
                    {
                        int excelRow = startRow + i - 1;
                        processedRows++;

                        string billTrackNumber = GetArrayValue(trackData, i, 1);
                        if (!string.IsNullOrWhiteSpace(billTrackNumber))
                        {
                            string normalizedTrack = NormalizeTrackNumber(billTrackNumber);
                            if (shippingIndex.ContainsKey(normalizedTrack))
                            {
                                List<ShippingItem> matchedItems = shippingIndex[normalizedTrack];
                                matchedCount++;

                                HashSet<string> productCodes = new HashSet<string>();
                                HashSet<string> productNames = new HashSet<string>();

                                foreach (ShippingItem item in matchedItems)
                                {
                                    if (!string.IsNullOrWhiteSpace(item.ProductCode))
                                        productCodes.Add(item.ProductCode.Trim());
                                    if (!string.IsNullOrWhiteSpace(item.ProductName))
                                        productNames.Add(item.ProductName.Trim());
                                }

                                if (productCodes.Count > 0)
                                    productUpdates[excelRow] = string.Join("、", productCodes.ToArray());
                                if (productNames.Count > 0)
                                    nameUpdates[excelRow] = string.Join("、", productNames.ToArray());
                            }
                        }
                    }

                    if (productUpdates.Count >= writeBatchSize || (endRow >= totalRows && productUpdates.Count > 0))
                    {
                        BatchUpdateColumn(billSheet, productCol, productUpdates);
                        result.UpdatedCells += productUpdates.Count;
                        productUpdates.Clear();
                    }
                    if (nameUpdates.Count >= writeBatchSize || (endRow >= totalRows && nameUpdates.Count > 0))
                    {
                        BatchUpdateColumn(billSheet, nameCol, nameUpdates);
                        result.UpdatedCells += nameUpdates.Count;
                        nameUpdates.Clear();
                    }

                    if (progressCallback != null)
                    {
                        int progress = 60 + (int)(35.0 * endRow / totalRows);
                        progressCallback(progress, string.Format("智能批处理账单: {0}/{1} 行 (已匹配: {2})", 
                            endRow, totalRows, matchedCount));
                    }
                }

                result.ProcessedRows = processedRows;
                result.MatchedCount = matchedCount;
            }
            catch (Exception ex)
            {
                WriteLog("平衡模式优化版账单处理错误: " + ex.ToString(), LogLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// 兼容模式发货明细索引构建 - 逐行读取
        /// </summary>
        private Dictionary<string, List<ShippingItem>> BuildShippingIndexCompatible(
            Excel.Worksheet shippingSheet, MultiWorkbookMatchConfig config, ProgressReportDelegate progressCallback)
        {
            Dictionary<string, List<ShippingItem>> index = new Dictionary<string, List<ShippingItem>>();

            try
            {
                Excel.Range usedRange = shippingSheet.UsedRange;
                if (usedRange == null || usedRange.Rows.Count < 2) return index;

                int totalRows = usedRange.Rows.Count;
                // --- FIX: Use the new method name from the user-provided ExcelHelper.
                int trackCol = ExcelHelper.GetColumnNumber(config.ShippingTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.ShippingProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.ShippingNameColumn);

                for (int row = 2; row <= totalRows; row++)
                {
                    string trackNumber = GetCellValue(shippingSheet, row, trackCol);
                    string productCode = GetCellValue(shippingSheet, row, productCol);
                    string productName = GetCellValue(shippingSheet, row, nameCol);

                    if (!string.IsNullOrWhiteSpace(trackNumber))
                    {
                        string normalizedTrack = NormalizeTrackNumber(trackNumber);
                        if (!index.ContainsKey(normalizedTrack))
                            index[normalizedTrack] = new List<ShippingItem>();

                        index[normalizedTrack].Add(new ShippingItem
                        {
                            ProductCode = productCode ?? "",
                            ProductName = productName ?? ""
                        });
                    }

                    if (row % 200 == 0 && progressCallback != null)
                    {
                        int progress = 15 + (int)(40.0 * row / totalRows);
                        progressCallback(progress, string.Format("构建索引: {0}/{1} 行", row, totalRows));
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("兼容模式索引构建错误: " + ex.ToString(), LogLevel.Error);
                throw;
            }

            return index;
        }

        /// <summary>
        /// 兼容模式账单明细处理 - 逐行处理
        /// </summary>
        private void ProcessBillDetailsCompatible(Excel.Worksheet billSheet, MultiWorkbookMatchConfig config,
            Dictionary<string, List<ShippingItem>> shippingIndex, MatchResult result, 
            ProgressReportDelegate progressCallback)
        {
            try
            {
                Excel.Range usedRange = billSheet.UsedRange;
                if (usedRange == null || usedRange.Rows.Count < 2) return;

                int totalRows = usedRange.Rows.Count;
                int matchedCount = 0;
                int processedRows = 0;
                int updatedCells = 0;
                
                // --- FIX: Use the new method name from the user-provided ExcelHelper.
                int trackCol = ExcelHelper.GetColumnNumber(config.BillTrackColumn);
                int productCol = ExcelHelper.GetColumnNumber(config.BillProductColumn);
                int nameCol = ExcelHelper.GetColumnNumber(config.BillNameColumn);

                for (int row = 2; row <= totalRows; row++)
                {
                    processedRows++;

                    string billTrackNumber = GetCellValue(billSheet, row, trackCol);
                    if (!string.IsNullOrWhiteSpace(billTrackNumber))
                    {
                        string normalizedTrack = NormalizeTrackNumber(billTrackNumber);
                        if (shippingIndex.ContainsKey(normalizedTrack))
                        {
                            List<ShippingItem> matchedItems = shippingIndex[normalizedTrack];
                            matchedCount++;

                            HashSet<string> productCodes = new HashSet<string>();
                            HashSet<string> productNames = new HashSet<string>();

                            foreach (ShippingItem item in matchedItems)
                            {
                                if (!string.IsNullOrWhiteSpace(item.ProductCode))
                                    productCodes.Add(item.ProductCode.Trim());
                                if (!string.IsNullOrWhiteSpace(item.ProductName))
                                    productNames.Add(item.ProductName.Trim());
                            }

                            if (productCodes.Count > 0)
                            {
                                SetCellValue(billSheet, row, productCol, 
                                    string.Join("、", productCodes.ToArray()));
                                updatedCells++;
                            }

                            if (productNames.Count > 0)
                            {
                                SetCellValue(billSheet, row, nameCol, 
                                    string.Join("、", productNames.ToArray()));
                                updatedCells++;
                            }
                        }
                    }

                    if (row % 100 == 0 && progressCallback != null)
                    {
                        int progress = 60 + (int)(35.0 * row / totalRows);
                        progressCallback(progress, string.Format("处理账单: {0}/{1} 行", row, totalRows));
                    }
                }

                result.ProcessedRows = processedRows;
                result.MatchedCount = matchedCount;
                result.UpdatedCells = updatedCells;
            }
            catch (Exception ex)
            {
                WriteLog("兼容模式账单处理错误: " + ex.ToString(), LogLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// 终极批量写入算法
        /// </summary>
        private int UltimateBatchWrite(Excel.Worksheet worksheet, int column, object[,] data, bool[] hasUpdate, int dataRows, int totalRows, string columnName)
        {
            try
            {
                if (data == null || dataRows == 0) return 0;

                WriteLog(string.Format("开始{0}列终极批量写入，数据行数: {1}", columnName, dataRows), LogLevel.Info);
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                string columnLetter = ExcelHelper.GetColumnLetter(column);
                
                Excel.Range targetRange = worksheet.Range[
                    string.Format("{0}2", columnLetter), 
                    string.Format("{0}{1}", columnLetter, totalRows)];

                targetRange.NumberFormat = "@";
                targetRange.Value2 = data;

                int updateCount = 0;
                for (int i = 0; i < dataRows; i++)
                {
                    if (hasUpdate[i]) updateCount++;
                }

                stopwatch.Stop();
                WriteLog(string.Format("{0}列终极批量写入完成 - 更新{1}个单元格，耗时{2:F3}秒", 
                    columnName, updateCount, stopwatch.Elapsed.TotalSeconds), LogLevel.Info);
                
                return updateCount;
            }
            catch (Exception ex)
            {
                WriteLog(string.Format("{0}列终极批量写入失败: {1}", columnName, ex.Message), LogLevel.Error);
                return 0;
            }
        }

        /// <summary>
        /// 批量更新列数据
        /// </summary>
        private int BatchUpdateColumn(Excel.Worksheet worksheet, int column, Dictionary<int, string> updates)
        {
            if (updates.Count == 0) return 0;

            try
            {
                var sortedUpdates = updates.OrderBy(x => x.Key).ToList();
                
                foreach (var update in sortedUpdates)
                {
                    Excel.Range cell = worksheet.Cells[update.Key, column] as Excel.Range;
                    if (cell != null)
                    {
                        cell.NumberFormat = "@";
                        cell.Value = update.Value;
                    }
                }

                WriteLog(string.Format("列{0}批量更新完成，共{1}个单元格", 
                    ExcelHelper.GetColumnLetter(column), updates.Count), LogLevel.Info);
                
                return updates.Count;
            }
            catch (Exception ex)
            {
                WriteLog(string.Format("批量更新列{0}失败: {1}", 
                    ExcelHelper.GetColumnLetter(column), ex.Message), LogLevel.Error);
                return 0;
            }
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        private string GetCellValue(Excel.Worksheet worksheet, int row, int column)
        {
            try
            {
                Excel.Range cell = worksheet.Cells[row, column] as Excel.Range;
                if (cell != null && cell.Value2 != null)
                {
                    return cell.Value2.ToString();
                }
                return "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        private void SetCellValue(Excel.Worksheet worksheet, int row, int column, string value)
        {
            try
            {
                Excel.Range cell = worksheet.Cells[row, column] as Excel.Range;
                if (cell != null)
                {
                    cell.NumberFormat = "@";
                    cell.Value2 = value;
                }
            }
            catch (Exception ex)
            {
                WriteLog("设置单元格值时出错：" + ex.ToString(), LogLevel.Error);
            }
        }

        /// <summary>
        /// 标准化运单号
        /// </summary>
        private string NormalizeTrackNumber(string trackNumber)
        {
            if (string.IsNullOrWhiteSpace(trackNumber))
                return "";

            string normalized = trackNumber.Trim();

            if (normalized.Contains("E+") || normalized.Contains("e+"))
            {
                if (decimal.TryParse(normalized, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out decimal decValue))
                {
                    normalized = decValue.ToString();
                }
            }

            return normalized;
        }

        /// <summary>
        /// 写入日志 (公共静态方法)
        /// </summary>
        public static void WriteLog(string message, LogLevel level)
        {
            try
            {
                if (!Directory.Exists(LogPath))
                {
                    Directory.CreateDirectory(LogPath);
                }

                string logFile = Path.Combine(LogPath, 
                    string.Format("YYTools_{0:yyyyMMdd}.log", DateTime.Now));

                string logEntry = string.Format("[{0:yyyy-MM-dd HH:mm:ss.fff}] [{1}] {2}",
                    DateTime.Now, level.ToString().ToUpper(), message);

                File.AppendAllText(logFile, logEntry + Environment.NewLine, Encoding.UTF8);

                System.Diagnostics.Debug.WriteLine(logEntry);
            }
            catch
            {
                // 日志写入失败不影响主流程
            }
        }

        /// <summary>
        /// 获取日志文件夹路径
        /// </summary>
        public static string GetLogFolderPath()
        {
            return LogPath;
        }

        /// <summary>
        /// 清理旧日志文件
        /// </summary>
        public static void CleanupOldLogs()
        {
            try
            {
                if (!Directory.Exists(LogPath)) return;

                var cutoffDate = DateTime.Now.AddDays(-7);
                var files = Directory.GetFiles(LogPath, "YYTools_*.log");

                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.CreationTime < cutoffDate)
                    {
                        fileInfo.Delete();
                    }
                }
            }
            catch
            {
                // 清理失败不影响主流程
            }
        }

        /// <summary>
        /// 安全获取数组值
        /// </summary>
        private string GetArrayValue(object[,] array, int row, int col)
        {
            try
            {
                if (array == null || row > array.GetLength(0) || col > array.GetLength(1) || row < 1 || col < 1)
                    return "";
                
                object value = array[row, col];
                return value != null ? value.ToString().Trim() : "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 计算最优批处理大小
        /// </summary>
        private int CalculateOptimalBatchSize()
        {
            try
            {
                int coreCount = Environment.ProcessorCount;
                long memoryMB = GC.GetTotalMemory(false) / 1024 / 1024;
                
                if (coreCount >= 8 && memoryMB > 4096)
                    return 2000; 
                else if (coreCount >= 4 && memoryMB > 2048)
                    return 1000; 
                else
                    return 500;
            }
            catch
            {
                return 1000;
            }
        }
    }

    /// <summary>
    /// 发货明细项
    /// </summary>
    public class ShippingItem
    {
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
    }

    /// <summary>
    /// 日志级别
    /// </summary>
    public enum LogLevel
    {
        Info,
        Warning,
        Error
    }
}