using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
            // 运单号相关列
            new SmartColumnRule(new[] { "运单", "快递单", "物流单", "tracking", "track", "单号" }, new[] { "TrackColumn" }, 10),
            new SmartColumnRule(new[] { "运单号", "快递单号", "物流单号" }, new[] { "TrackColumn" }, 9),
            
            // 商品编码相关列
            new SmartColumnRule(new[] { "商品编码", "产品编码", "sku", "编码" }, new[] { "ProductColumn" }, 10),
            new SmartColumnRule(new[] { "商品", "产品", "product" }, new[] { "ProductColumn" }, 8),
            
            // 商品名称相关列
            new SmartColumnRule(new[] { "商品名称", "产品名称", "品名", "名称" }, new[] { "NameColumn" }, 10),
            new SmartColumnRule(new[] { "商品", "产品", "product" }, new[] { "NameColumn" }, 7),
            
            // 数量相关列
            new SmartColumnRule(new[] { "数量", "件数", "qty", "quantity" }, new[] { "QuantityColumn" }, 9),
            
            // 价格相关列
            new SmartColumnRule(new[] { "价格", "单价", "金额", "price", "amount" }, new[] { "PriceColumn" }, 9),
        };

        /// <summary>
        /// 获取工作表的列信息
        /// </summary>
        public static List<ColumnInfo> GetColumnInfos(Excel.Worksheet worksheet, int maxRowsForPreview = 100)
        {
            var columns = new List<ColumnInfo>();
            
            try
            {
                if (worksheet == null) return columns;

                var usedRange = worksheet.UsedRange;
                if (usedRange.Rows.Count == 0) return columns;

                int colCount = usedRange.Columns.Count;
                int rowCount = Math.Min(usedRange.Rows.Count, maxRowsForPreview);

                for (int i = 1; i <= colCount; i++)
                {
                    string colLetter = ExcelHelper.GetColumnLetter(i);
                    // 按列扫描：第一条非空作为标题，第二条非空作为预览
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
                            // 前5行都为空，直接放弃深入解析，避免卡顿
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
            
            try
            {
                foreach (var rule in ColumnRules.OrderByDescending(r => r.Priority))
                {
                    foreach (var columnType in rule.ColumnTypes)
                    {
                        if (matchedColumns.ContainsKey(columnType)) continue;

                        var bestMatch = FindBestMatch(columns, rule.Keywords);
                        if (bestMatch != null)
                        {
                            matchedColumns[columnType] = bestMatch;
                            break;
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

        /// <summary>
        /// 搜索列
        /// </summary>
        public static List<ColumnInfo> SearchColumns(List<ColumnInfo> columns, string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return columns;

            try
            {
                var searchTerms = ParseSearchText(searchText);
                return columns.Where(col => 
                    searchTerms.Any(term => 
                        col.HeaderText?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        col.PreviewData?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        col.ColumnLetter?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0
                    )).ToList();
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"列搜索失败: {ex.Message}", LogLevel.Error);
                return columns;
            }
        }

        /// <summary>
        /// 验证列选择
        /// </summary>
        public static bool ValidateColumnSelection(ColumnInfo column, string expectedType)
        {
            if (column == null || !column.IsValid) return false;

            // 根据类型进行验证
            switch (expectedType)
            {
                case "TrackColumn":
                    return !string.IsNullOrWhiteSpace(column.HeaderText) && 
                           (column.HeaderText.Contains("运单") || column.HeaderText.Contains("快递") || 
                            column.HeaderText.Contains("物流") || column.HeaderText.Contains("单号"));
                
                case "ProductColumn":
                    return !string.IsNullOrWhiteSpace(column.HeaderText) && 
                           (column.HeaderText.Contains("编码") || column.HeaderText.Contains("商品") || 
                            column.HeaderText.Contains("产品"));
                
                case "NameColumn":
                    return !string.IsNullOrWhiteSpace(column.HeaderText) && 
                           (column.HeaderText.Contains("名称") || column.HeaderText.Contains("品名") || 
                            column.HeaderText.Contains("商品") || column.HeaderText.Contains("产品"));
                
                default:
                    return column.IsValid;
            }
        }

        #region 私有方法

        private static Excel.Range FindHeaderRow(Excel.Range usedRange)
        {
            for (int i = 1; i <= Math.Min(100, usedRange.Rows.Count); i++)
            {
                var row = usedRange.Rows[i] as Excel.Range;
                var rowData = row.Value2 as object[,];
                if (rowData != null && Enumerable.Range(1, rowData.GetLength(1)).Any(col => 
                    rowData[1, col] != null && !string.IsNullOrWhiteSpace(rowData[1, col].ToString())))
                {
                    return row;
                }
            }
            return usedRange.Rows[1] as Excel.Range;
        }

        private static string GetHeaderText(object[,] headers, int columnIndex)
        {
            try
            {
                if (headers != null && columnIndex <= headers.GetLength(1))
                {
                    return headers[1, columnIndex]?.ToString().Trim() ?? "";
                }
            }
            catch { }
            return "";
        }

        private static string GetPreviewData(Excel.Worksheet worksheet, int columnIndex, Excel.Range headerRow, int maxRows)
        {
            try
            {
                if (headerRow == null) return "";

                // 从标题行下一行开始查找非空数据
                for (int row = headerRow.Row + 1; row <= Math.Min(headerRow.Row + maxRows, worksheet.UsedRange.Rows.Count); row++)
                {
                    var cell = worksheet.Cells[row, columnIndex] as Excel.Range;
                    if (cell != null)
                    {
                        var value = cell.Value2?.ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            return value.Length > 50 ? value.Substring(0, 50) + "..." : value;
                        }
                    }
                }
            }
            catch { }
            return "";
        }

        private static ColumnInfo FindBestMatch(List<ColumnInfo> columns, string[] keywords)
        {
            var bestMatch = columns
                .Where(col => col.IsValid && !string.IsNullOrWhiteSpace(col.HeaderText))
                .Select(col => new
                {
                    Column = col,
                    Score = CalculateMatchScore(col.HeaderText, keywords)
                })
                .Where(match => match.Score > 0)
                .OrderByDescending(match => match.Score)
                .FirstOrDefault();

            return bestMatch?.Column;
        }

        private static int CalculateMatchScore(string text, string[] keywords)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;

            int score = 0;
            string lowerText = text.ToLowerInvariant();

            foreach (var keyword in keywords)
            {
                if (lowerText.Contains(keyword.ToLowerInvariant()))
                {
                    score += keyword.Length;
                    // 完全匹配加分
                    if (lowerText.Equals(keyword.ToLowerInvariant()))
                        score += 10;
                }
            }

            return score;
        }

        private static string[] ParseSearchText(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText)) return new string[0];

            // 支持多种分隔符：空格、逗号、分号、括号等
            var separators = new[] { ' ', ',', ';', '(', ')', '（', '）', '、' };
            return searchText.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                           .Where(s => !string.IsNullOrWhiteSpace(s))
                           .Select(s => s.Trim())
                           .ToArray();
        }

        #endregion
    }
}