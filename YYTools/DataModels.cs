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
    /// 日志级别
    /// </summary>
    public enum LogLevel
    {
        Info,
        Warning,
        Error
    }
    
    /// <summary>
    /// 列信息类
    /// </summary>
    public class ColumnInfo
    {
        public string DisplayText { get; set; }
        public string ColumnLetter { get; set; }
        public string HeaderText { get; set; }
        public string PreviewData { get; set; }
        public string SearchKeywords { get; set; }
    }
}