using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// Excel辅助工具（从旧版移植）
    /// </summary>
    public static class ExcelHelper
    {
        public static string GetColumnLetter(int columnNumber)
        {
            if (columnNumber <= 0) throw new ArgumentException("列号必须大于0");
            string columnLetter = string.Empty;
            while (columnNumber > 0)
            {
                columnNumber--;
                columnLetter = (char)('A' + columnNumber % 26) + columnLetter;
                columnNumber /= 26;
            }
            return columnLetter;
        }

        public static int GetColumnNumber(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter)) throw new ArgumentException("列字母不能为空");
            columnLetter = columnLetter.ToUpperInvariant();
            int columnNumber = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                char letter = columnLetter[i];
                if (letter < 'A' || letter > 'Z') throw new ArgumentException($"无效的列字母：{letter}");
                columnNumber = columnNumber * 26 + (letter - 'A' + 1);
            }
            return columnNumber;
        }

        public static bool IsValidColumnLetter(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter)) return false;
            return columnLetter.ToUpperInvariant().All(c => c >= 'A' && c <= 'Z');
        }

        public static string GetCellValue(Excel.Range cell)
        {
            try { return cell?.Value2?.ToString().Trim() ?? string.Empty; } catch { return string.Empty; }
        }

        public static List<string> GetColumnDataBatch(Excel.Worksheet worksheet, string columnLetter, int startRow, int endRow)
        {
            var data = new List<string>();
            try
            {
                if (worksheet == null || string.IsNullOrWhiteSpace(columnLetter)) return data;
                var range = worksheet.Range[$"{columnLetter}{startRow}:{columnLetter}{endRow}"];
                var values = range?.Value2 as object[,];
                if (values != null)
                {
                    for (int i = 1; i <= values.GetLength(0); i++)
                    {
                        data.Add(values[i, 1]?.ToString().Trim() ?? string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"批量获取列数据失败: {ex.Message}");
            }
            return data;
        }

        public static (int rows, int columns, long dataSize) GetWorksheetStats(Excel.Worksheet worksheet)
        {
            try
            {
                if (worksheet == null) return (0, 0, 0);
                var usedRange = worksheet.UsedRange;
                int rows = usedRange.Rows.Count;
                int columns = usedRange.Columns.Count;
                long dataSize = (long)rows * columns * 50;
                return (rows, columns, dataSize);
            }
            catch (Exception ex)
            {
                Logger.LogError($"获取工作表统计信息失败: {ex.Message}");
                return (0, 0, 0);
            }
        }

        public static void OptimizeExcelPerformance(Excel.Application excelApp)
        {
            if (excelApp == null) return;
            try { excelApp.ScreenUpdating = false; } catch { }
            try { excelApp.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { }
            try { excelApp.EnableEvents = false; } catch { }
            try { excelApp.DisplayStatusBar = false; } catch { }
            try { excelApp.DisplayAlerts = false; } catch { }
        }

        public static void RestoreExcelPerformance(Excel.Application excelApp, bool originalScreenUpdating, Excel.XlCalculation originalCalculation, bool originalEnableEvents, bool originalDisplayStatusBar, bool originalDisplayAlerts)
        {
            if (excelApp == null) return;
            try { excelApp.ScreenUpdating = originalScreenUpdating; } catch { }
            try { excelApp.Calculation = originalCalculation; } catch { }
            try { excelApp.EnableEvents = originalEnableEvents; } catch { }
            try { excelApp.DisplayStatusBar = originalDisplayStatusBar; } catch { }
            try { excelApp.DisplayAlerts = originalDisplayAlerts; } catch { }
        }
    }
}

