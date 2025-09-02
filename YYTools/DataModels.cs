// --- 文件 1: DataModels.cs ---

using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
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
        public SortOption SortOption { get; set; }

        // 添加拼接和去重配置
        public string ConcatenationDelimiter { get; set; } = "、";
        public bool RemoveDuplicateItems { get; set; } = true;
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

    public enum SortOption
    {
        None = 0,
        Asc = 1,
        Desc = 2
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
    /// 工作簿信息类 (最终定义位置)
    /// </summary>
    public class WorkbookInfo
    {
        public string Name { get; set; }
        public Excel.Workbook Workbook { get; set; }
        public bool IsActive { get; set; }
    }

    /// <summary>
    /// 列信息类
    /// </summary>
    public class ColumnInfo
    {
        public string ColumnLetter { get; set; }
        public string HeaderText { get; set; }
        public string PreviewData { get; set; } // 此字段在禁用预览时可能为空
        public int RowCount { get; set; }
        public bool IsValid { get; set; }

        public override string ToString()
        {
            // 如果预览数据为空 (因为功能被禁用或确实没有数据)，则显示简化模式
            if (string.IsNullOrWhiteSpace(PreviewData))
            {
                return string.IsNullOrWhiteSpace(HeaderText)
                    ? $"{ColumnLetter}"
                    : $"{ColumnLetter}: {HeaderText}";
            }

            // 完整预览模式
            return $"{ColumnLetter}: {HeaderText} (示例: {PreviewData})";
        }
    }

    /// <summary>
    /// 智能列匹配规则
    /// </summary>
    public class SmartColumnRule
    {
        public string[] Keywords { get; set; }
        public string[] ColumnTypes { get; set; }
        public int Priority { get; set; }

        public SmartColumnRule(string[] keywords, string[] columnTypes, int priority = 1)
        {
            Keywords = keywords;
            ColumnTypes = columnTypes;
            Priority = priority;
        }
    }

    /// <summary>
    /// 日志级别
    /// </summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    /// <summary>
    /// 账单行数据模型，用于预处理和缓存
    /// </summary>
    public class BillRowData
    {
        /// <summary>
        /// 运单号
        /// </summary>
        public string TrackNumber { get; set; }

        /// <summary>
        /// 商品编码列号
        /// </summary>
        public int ProductColumn { get; set; }

        /// <summary>
        /// 商品名称列号
        /// </summary>
        public int NameColumn { get; set; }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowNumber { get; set; }
    }
}