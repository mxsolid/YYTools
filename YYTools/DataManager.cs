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
        // 针对列解析的键级别锁，防止相同工作表被并发重复解析
        private static readonly Dictionary<string, object> columnKeyLocks = new Dictionary<string, object>();
        // 并发限制信号量，防止过多并发导致内存与句柄压力
        private static System.Threading.SemaphoreSlim parseSemaphore = new System.Threading.SemaphoreSlim(System.Environment.ProcessorCount);

        /// <summary>
        /// 更新解析的最大并发数（从设置中读取），上限不超过 CPU 核心数
        /// </summary>
        public static void UpdateMaxConcurrency(int requested)
        {
            int limit = System.Math.Max(1, System.Math.Min(requested, System.Environment.ProcessorCount));
            // 直接替换信号量实例，旧实例允许自然释放
            parseSemaphore = new System.Threading.SemaphoreSlim(limit);
            Logger.LogInfo($"更新列解析最大并发为: {limit}");
        }

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
            // 先查缓存
            lock (cacheLock)
            {
                if (columnInfoCache.TryGetValue(key, out var cached))
                {
                    Logger.LogInfo($"从缓存命中列信息: {worksheet.Name}");
                    return cached;
                }
            }

            // 并发限制，避免内存与COM对象压力
            parseSemaphore.Wait();
            try
            {
                // 双重检查缓存，防止并发重复解析
                lock (cacheLock)
                {
                    if (columnInfoCache.TryGetValue(key, out var cached2))
                    {
                        Logger.LogInfo($"从缓存命中列信息: {worksheet.Name}");
                        return cached2;
                    }
                }

                // 针对该key的局部锁
                object localLock;
                lock (cacheLock)
                {
                    if (!columnKeyLocks.TryGetValue(key, out localLock))
                    {
                        localLock = new object();
                        columnKeyLocks[key] = localLock;
                    }
                }

                lock (localLock)
                {
                    // 再次检查缓存
                    lock (cacheLock)
                    {
                        if (columnInfoCache.TryGetValue(key, out var cached3))
                        {
                            Logger.LogInfo($"从缓存命中列信息: {worksheet.Name}");
                            return cached3;
                        }
                    }

                    Logger.LogInfo($"从Excel读取列信息: {worksheet?.Name}");
                    var infos = SmartColumnService.GetColumnInfos(worksheet, 20);
                    lock (cacheLock)
                    {
                        columnInfoCache[key] = infos;
                    }
                    return infos;
                }
            }
            finally
            {
                parseSemaphore.Release();
            }
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