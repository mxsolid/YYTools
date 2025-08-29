using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// 智能列选择服务
    /// </summary>
    public class SmartColumnService
    {
        private static readonly List<SmartColumnRule> ColumnRules = new List<SmartColumnRule>
        {
            // 运单号相关列 (新规则，高优先级)
            new SmartColumnRule(new[] { "快递单号", "运单号", "邮件号", "物流单号", "快递号" , "物流号" }, new[] { "TrackColumn" }, 10),
            new SmartColumnRule(new[] { "运单", "快递", "物流", "单号", "tracking" }, new[] { "TrackColumn" }, 8),
            
            // 商品名称相关列 (新规则，高优先级)
            new SmartColumnRule(new[] { "商品名称", "品名" }, new[] { "NameColumn" }, 10),
            new SmartColumnRule(new[] { "名称" }, new[] { "NameColumn" }, 7),
            
            // 商品编码相关列 (新规则，高优先级)
            // "商品"作为完全匹配项，优先级适中，避免错误匹配"商品名称"
            new SmartColumnRule(new[] { "商品" }, new[] { "ProductColumn" }, 10),
            new SmartColumnRule(new[] { "商品编码", "商品编号", "产品编码", "商品代码", "款号", "sku" }, new[] { "ProductColumn" }, 8),
            new SmartColumnRule(new[] { "编码" }, new[] { "ProductColumn" }, 6),

            
            // 其他通用规则，优先级较低
            new SmartColumnRule(new[] { "产品" }, new[] { "ProductColumn", "NameColumn" }, 5),
            new SmartColumnRule(new[] { "数量", "件数", "qty", "quantity" }, new[] { "QuantityColumn" }, 9),
            new SmartColumnRule(new[] { "价格", "单价", "金额", "price", "amount" }, new[] { "PriceColumn" }, 9),
        };

        /// <summary>
        /// 获取工作表的列信息（性能优化版本）
        /// </summary>
        /// <param name="worksheet">要解析的工作表</param>
        /// <param name="enableDataPreview">是否启用数据预览</param>
        public static List<ColumnInfo> GetColumnInfos(Excel.Worksheet worksheet, int maxRowsForPreview, bool enableDataPreview)
        {
            var columns = new List<ColumnInfo>();
            try
            {
                if (worksheet == null) return columns;

                var usedRange = worksheet.UsedRange;
                if (usedRange.Rows.Count == 0) return columns;
                
                int colCount = Math.Min(usedRange.Columns.Count, 256); // 限制最大扫描列数
                
                // 根据是否启用预览，决定扫描深度
                int headerScanDepth = 3; // 最多扫描3行来找标题
                int dataPreviewScanDepth = maxRowsForPreview; // 扫描20行来找预览数据

                for (int i = 1; i <= colCount; i++)
                {
                    string colLetter = ExcelHelper.GetColumnLetter(i);
                    string headerText = "";
                    string previewData = "";

                    // 步骤1: 查找标题 (最多扫描3行)
                    for (int row = 1; row <= headerScanDepth; row++)
                    {
                        if (row > usedRange.Rows.Count) break;
                        var cellValue = (worksheet.Cells[row, i] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            headerText = cellValue;
                            break; // 找到第一个非空单元格作为标题
                        }
                    }

                    // 步骤2: 如果启用预览，则查找预览数据
                    if (enableDataPreview && usedRange.Rows.Count > 1)
                    {
                        for (int row = 2; row <= dataPreviewScanDepth; row++) // 从第二行开始找数据
                        {
                            if (row > usedRange.Rows.Count) break;
                            var cellValue = (worksheet.Cells[row, i] as Excel.Range)?.Value2?.ToString().Trim() ?? "";
                            // 确保预览数据和标题不一样
                            if (!string.IsNullOrWhiteSpace(cellValue) && cellValue != headerText)
                            {
                                previewData = cellValue.Length > 50 ? cellValue.Substring(0, 50) + "..." : cellValue;
                                break;
                            }
                        }
                    }

                    columns.Add(new ColumnInfo
                    {
                        ColumnLetter = colLetter,
                        HeaderText = headerText,
                        PreviewData = previewData, // 如果未启用预览，此项为空
                        IsValid = !string.IsNullOrWhiteSpace(headerText) || !string.IsNullOrWhiteSpace(previewData)
                    });
                }
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"获取列信息失败: {ex.Message}", LogLevel.Error);
            }

            return columns;
        }
    

        /// <summary>
        /// 智能匹配列
        /// </summary>
        public static Dictionary<string, ColumnInfo> SmartMatchColumns(List<ColumnInfo> columns)
        {
            var matchedColumns = new Dictionary<string, ColumnInfo>();
            var usedColumns = new HashSet<ColumnInfo>();
            
            try
            {
                foreach (var rule in ColumnRules.OrderByDescending(r => r.Priority))
                {
                    foreach (var columnType in rule.ColumnTypes)
                    {
                        if (matchedColumns.ContainsKey(columnType)) continue;

                        var bestMatch = FindBestMatch(columns.Except(usedColumns).ToList(), rule.Keywords, rule.Keywords.First() == "商品");
                        if (bestMatch != null)
                        {
                            matchedColumns[columnType] = bestMatch;
                            usedColumns.Add(bestMatch);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"智能列匹配失败: {ex.Message}", LogLevel.Error);
            }

            return matchedColumns;
        }

        private static ColumnInfo FindBestMatch(List<ColumnInfo> columns, string[] keywords, bool exactMatchOnly = false)
        {
            var bestMatch = columns
                .Where(col => col.IsValid && !string.IsNullOrWhiteSpace(col.HeaderText))
                .Select(col => new
                {
                    Column = col,
                    Score = CalculateMatchScore(col.HeaderText, keywords, exactMatchOnly)
                })
                .Where(match => match.Score > 0)
                .OrderByDescending(match => match.Score)
                .FirstOrDefault();

            return bestMatch?.Column;
        }

        private static int CalculateMatchScore(string text, string[] keywords, bool exactMatchOnly)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;

            int score = 0;
            string lowerText = text.Trim().ToLowerInvariant();

            foreach (var keyword in keywords)
            {
                string lowerKeyword = keyword.ToLowerInvariant();
                
                // 完全匹配模式
                if (exactMatchOnly)
                {
                    if (lowerText.Equals(lowerKeyword))
                    {
                        score = 100; // Give a very high score for exact match
                        break;
                    }
                }
                // 包含匹配模式
                else
                {
                    if (lowerText.Contains(lowerKeyword))
                    {
                        score += keyword.Length;
                        if (lowerText.Equals(lowerKeyword))
                            score += 10;
                    }
                }
            }

            return score;
        }
    }
}