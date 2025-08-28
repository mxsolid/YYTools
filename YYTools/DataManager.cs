using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    public static class DataManager
    {
        private static readonly Dictionary<int, List<string>> sheetNamesCache = new Dictionary<int, List<string>>();
        private static readonly Dictionary<int, List<ColumnInfo>> columnInfoCache = new Dictionary<int, List<ColumnInfo>>();

        public static List<string> GetSheetNames(Excel.Workbook workbook)
        {
            int hashCode = workbook.GetHashCode();
            if (sheetNamesCache.ContainsKey(hashCode))
            {
                Logger.LogInfo($"从缓存命中工作表列表: {workbook.Name}");
                return sheetNamesCache[hashCode];
            }

            Logger.LogInfo($"从Excel读取工作表列表: {workbook.Name}");
            var names = ExcelAddin.GetWorksheetNames(workbook);
            sheetNamesCache[hashCode] = names;
            return names;
        }

        public static List<ColumnInfo> GetColumnInfos(Excel.Worksheet worksheet)
        {
            int hashCode = worksheet.GetHashCode();
            if (columnInfoCache.ContainsKey(hashCode))
            {
                Logger.LogInfo($"从缓存命中列信息: {worksheet.Name}");
                return columnInfoCache[hashCode];
            }

            Logger.LogInfo($"从Excel读取列信息: {worksheet.Name}");
            var infos = SmartColumnService.GetColumnInfos(worksheet, 50);
            columnInfoCache[hashCode] = infos;
            return infos;
        }

        public static void ClearCache()
        {
            sheetNamesCache.Clear();
            columnInfoCache.Clear();
            Logger.LogInfo("所有数据缓存已清空。");
        }
    }
}