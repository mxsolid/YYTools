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
        public string TrackNumber { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
    }

    /// <summary>
    /// 账单明细项
    /// </summary>
    public class BillItem
    {
        public string TrackNumber { get; set; }
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
        public string PreviewData { get; set; }
        public int RowCount { get; set; }
        public bool IsValid { get; set; }
        
        public override string ToString()
        {
            string title = string.IsNullOrWhiteSpace(HeaderText) ? "" : HeaderText;
            string preview = string.IsNullOrWhiteSpace(PreviewData) ? "" : $" | 示例: {PreviewData}";
            return string.IsNullOrWhiteSpace(title)
                ? ColumnLetter
                : $"{ColumnLetter} ({title}){preview}";
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
}