using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    public static class DataManager
    {
        // 使用稳定键(Workbook.FullName + Worksheet.Name)并加锁，避免哈希冲突和线程不安全
        private static readonly object cacheLock = new object();
        private static readonly Dictionary<string, List<string>> sheetNamesCache = new Dictionary<string, List<string>>();
        private static readonly Dictionary<string, List<ColumnInfo>> columnInfoCache = new Dictionary<string, List<ColumnInfo>>();

        public static List<string> GetSheetNames(Excel.Workbook workbook)
        {
            string key = (workbook != null ? (workbook.FullName ?? workbook.Name) : "") + "::sheets";
            lock (cacheLock)
            {
                if (sheetNamesCache.TryGetValue(key, out var cached))
                {
                    Logger.LogInfo($"从缓存命中工作表列表: {workbook.Name}");
                    return cached;
                }
            }

            Logger.LogInfo($"从Excel读取工作表列表: {workbook.Name}");
            var names = ExcelAddin.GetWorksheetNames(workbook);
            lock (cacheLock)
            {
                sheetNamesCache[key] = names;
            }
            return names;
        }

        public static List<ColumnInfo> GetColumnInfos(Excel.Worksheet worksheet)
        {
            string wbName = worksheet?.Parent is Excel.Workbook wb ? (wb.FullName ?? wb.Name) : "";
            string key = wbName + "::" + (worksheet?.Name ?? "") + "::columns";
            lock (cacheLock)
            {
                if (columnInfoCache.TryGetValue(key, out var cached))
                {
                    Logger.LogInfo($"从缓存命中列信息: {worksheet.Name}");
                    return cached;
                }
            }

            Logger.LogInfo($"从Excel读取列信息: {worksheet.Name}");
            var infos = SmartColumnService.GetColumnInfos(worksheet, 50);
            lock (cacheLock)
            {
                columnInfoCache[key] = infos;
            }
            return infos;
        }

        public static void ClearCache()
        {
            lock (cacheLock)
            {
                sheetNamesCache.Clear();
                columnInfoCache.Clear();
            }
            Logger.LogInfo("所有数据缓存已清空。");
        }
    }
}