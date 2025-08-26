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
}