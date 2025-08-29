using System;
using System.Collections.Generic;
using System.Linq;
using YYTools.Pricing;
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
            new SmartColumnRule(PricingKeywords.QuantityKeywords, new[] { "QuantityColumn" }, 9),
            new SmartColumnRule(PricingKeywords.PriceKeywords, new[] { "PriceColumn" }, 9),
        };

        /// <summary>
        /// 获取工作表的列信息（性能优化版本）
        /// </summary>
        public static List<ColumnInfo> GetColumnInfos(Excel.Worksheet worksheet, int maxRowsForPreview = 10)
        {
            var columns = new List<ColumnInfo>();
            
            try
            {
                if (worksheet == null) return columns;

                // 性能优化：减少锁的使用，提高性能
                var usedRange = worksheet.UsedRange;
                if (usedRange.Rows.Count == 0) return columns;

                int colCount = Math.Min(usedRange.Columns.Count, 100); // 限制最大扫描列数防止卡顿
                int rowCount = Math.Min(usedRange.Rows.Count, maxRowsForPreview);

                for (int i = 1; i <= colCount; i++)
                {
                    string colLetter = ExcelHelper.GetColumnLetter(i);
                    string headerText = "";
                    string previewData = "";
                    int found = 0;
                    int maxScan = Math.Min(usedRange.Rows.Count, rowCount);
                    bool firstFiveAllEmpty = true;
                    for (int row = 1; row <= maxScan; row++)
                    {
                        var cell = worksheet.Cells[row, i] as Excel.Range;
                        string value = cell?.Value2?.ToString().Trim() ?? "";
                        if (row <= 5 && !string.IsNullOrWhiteSpace(value)) firstFiveAllEmpty = false;
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            found++;
                            if (found == 1) headerText = value;
                            else if (found == 2) { previewData = value.Length > 50 ? value.Substring(0, 50) + "..." : value; break; }
                        }
                        if (row == 5 && firstFiveAllEmpty)
                        {
                            break;
                        }
                    }
                    bool isValid = !string.IsNullOrWhiteSpace(headerText) || !string.IsNullOrWhiteSpace(previewData);

                    columns.Add(new ColumnInfo
                    {
                        ColumnLetter = colLetter,
                        HeaderText = headerText,
                        PreviewData = previewData,
                        RowCount = usedRange.Rows.Count,
                        IsValid = isValid
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