using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// 匹配配置类（移植）
    /// </summary>
    public class MatchConfig
    {
        public string ShippingSheetName { get; set; } = string.Empty;
        public string BillSheetName { get; set; } = string.Empty;
        public string ShippingTrackColumn { get; set; } = string.Empty;
        public string ShippingProductColumn { get; set; } = string.Empty;
        public string ShippingNameColumn { get; set; } = string.Empty;
        public string BillTrackColumn { get; set; } = string.Empty;
        public string BillProductColumn { get; set; } = string.Empty;
        public string BillNameColumn { get; set; } = string.Empty;
    }

    public class MultiWorkbookMatchConfig : MatchConfig
    {
        public Excel.Workbook ShippingWorkbook { get; set; }
        public Excel.Workbook BillWorkbook { get; set; }
        public SortOption SortOption { get; set; }
        public string ConcatenationDelimiter { get; set; } = "、";
        public bool RemoveDuplicateItems { get; set; } = true;
    }

    public class MatchResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
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

    public class ShippingItem
    {
        public string? ProductCode { get; set; }
        public string? ProductName { get; set; }
    }

    public class WorkbookInfo
    {
        public string Name { get; set; } = string.Empty;
        public Excel.Workbook Workbook { get; set; }
        public bool IsActive { get; set; }
    }

    public class ColumnInfo
    {
        public string ColumnLetter { get; set; } = string.Empty;
        public string HeaderText { get; set; } = string.Empty;
        public string? PreviewData { get; set; }
        public int RowCount { get; set; }
        public bool IsValid { get; set; }

        public override string ToString()
        {
            if (string.IsNullOrWhiteSpace(PreviewData))
            {
                return string.IsNullOrWhiteSpace(HeaderText)
                    ? $"{ColumnLetter}"
                    : $"{ColumnLetter}: {HeaderText}";
            }
            return $"{ColumnLetter}: {HeaderText} (示例: {PreviewData})";
        }
    }

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

    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    public class BillRowData
    {
        public string TrackNumber { get; set; } = string.Empty;
        public int ProductColumn { get; set; }
        public int NameColumn { get; set; }
        public int RowNumber { get; set; }
    }
}

