// --- 文件 4: ExcelHelper.cs ---
using System;
using System.Linq;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    public static class ExcelHelper
    {
        public static string GetColumnLetter(int columnNumber)
        {
            if (columnNumber <= 0)
                throw new ArgumentException("列号必须大于0");

            string columnLetter = "";
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
            if (string.IsNullOrEmpty(columnLetter))
                throw new ArgumentException("列字母不能为空");

            columnLetter = columnLetter.ToUpper();
            int columnNumber = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                char letter = columnLetter[i];
                if (letter < 'A' || letter > 'Z')
                    throw new ArgumentException($"无效的列字母：{letter}");

                columnNumber = columnNumber * 26 + (letter - 'A' + 1);
            }

            return columnNumber;
        }
        
        public static bool IsValidColumnLetter(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter))
                return false;
            return columnLetter.ToUpper().All(c => c >= 'A' && c <= 'Z');
        }

        /// <summary>
        /// 安全获取单元格值
        /// </summary>
        public static string GetCellValue(Excel.Range cell)
        {
            try
            {
                if (cell == null) return "";
                var value = cell.Value2;
                return value?.ToString().Trim() ?? "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 批量获取列数据（性能优化）
        /// </summary>
        public static List<string> GetColumnDataBatch(Excel.Worksheet worksheet, string columnLetter, int startRow, int endRow)
        {
            var data = new List<string>();
            try
            {
                if (worksheet == null || string.IsNullOrWhiteSpace(columnLetter)) return data;

                var range = worksheet.Range[$"{columnLetter}{startRow}:{columnLetter}{endRow}"];
                if (range == null) return data;

                var values = range.Value2 as object[,];
                if (values != null)
                {
                    for (int i = 1; i <= values.GetLength(0); i++)
                    {
                        var value = values[i, 1]?.ToString().Trim() ?? "";
                        data.Add(value);
                    }
                }
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"批量获取列数据失败: {ex.Message}", LogLevel.Error);
            }
            return data;
        }

        /// <summary>
        /// 安全设置单元格值
        /// </summary>
        public static bool SetCellValue(Excel.Range cell, string value)
        {
            try
            {
                if (cell == null) return false;
                cell.Value2 = value;
                return true;
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"设置单元格值失败: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// 获取工作表的统计信息
        /// </summary>
        public static (int rows, int columns, long dataSize) GetWorksheetStats(Excel.Worksheet worksheet)
        {
            try
            {
                if (worksheet == null) return (0, 0, 0);

                var usedRange = worksheet.UsedRange;
                if (usedRange == null) return (0, 0, 0);

                int rows = usedRange.Rows.Count;
                int columns = usedRange.Columns.Count;
                
                // 估算数据大小（粗略计算）
                long dataSize = (long)rows * columns * 50; // 假设每个单元格平均50字节

                return (rows, columns, dataSize);
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"获取工作表统计信息失败: {ex.Message}", LogLevel.Error);
                return (0, 0, 0);
            }
        }

        /// <summary>
        /// 检查工作表是否过大（性能警告）
        /// </summary>
        public static bool IsWorksheetTooLarge(Excel.Worksheet worksheet, int maxRows = 100000, int maxColumns = 1000)
        {
            try
            {
                var stats = GetWorksheetStats(worksheet);
                return stats.rows > maxRows || stats.columns > maxColumns;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 优化Excel应用程序性能
        /// </summary>
        public static void OptimizeExcelPerformance(Excel.Application excelApp)
        {
            try
            {
                if (excelApp == null) return;

                excelApp.ScreenUpdating = false;
                excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                excelApp.EnableEvents = false;
                excelApp.DisplayStatusBar = false;
                excelApp.DisplayAlerts = false;
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"优化Excel性能失败: {ex.Message}", LogLevel.Warning);
            }
        }

        /// <summary>
        /// 恢复Excel应用程序性能设置
        /// </summary>
        public static void RestoreExcelPerformance(Excel.Application excelApp, bool originalScreenUpdating, Excel.XlCalculation originalCalculation)
        {
            try
            {
                if (excelApp == null) return;

                excelApp.ScreenUpdating = originalScreenUpdating;
                excelApp.Calculation = originalCalculation;
                excelApp.EnableEvents = true;
                excelApp.DisplayStatusBar = true;
                excelApp.DisplayAlerts = true;
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"恢复Excel性能设置失败: {ex.Message}", LogLevel.Warning);
            }
        }
    }
}