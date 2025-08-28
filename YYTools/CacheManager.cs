using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// 高性能缓存管理器
    /// </summary>
    public class CacheManager : IDisposable
    {
        private static readonly Lazy<CacheManager> _instance = new Lazy<CacheManager>(() => new CacheManager());
        public static CacheManager Instance => _instance.Value;

        // 缓存存储
        private readonly ConcurrentDictionary<string, CachedWorkbook> _workbookCache = new ConcurrentDictionary<string, CachedWorkbook>();
        private readonly ConcurrentDictionary<string, CachedWorksheet> _worksheetCache = new ConcurrentDictionary<string, CachedWorksheet>();
        private readonly ConcurrentDictionary<string, CachedColumnInfo> _columnCache = new ConcurrentDictionary<string, CachedColumnInfo>();
        
        // 缓存配置
        private readonly int _maxWorkbooks = Constants.MaxCachedWorkbooks;
        private readonly int _maxWorksheets = Constants.MaxCachedWorksheets;
        private readonly int _maxColumns = Constants.MaxCachedColumns;
        private readonly TimeSpan _defaultExpiration = TimeSpan.FromMinutes(Constants.DefaultCacheExpirationMinutes);
        
        // 清理任务
        private readonly Timer _cleanupTimer;
        private bool _disposed = false;

        private CacheManager()
        {
            // 每5分钟清理一次过期缓存
            _cleanupTimer = new Timer(CleanupExpiredCache, null, TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
            
            Logger.LogInfo("缓存管理器已初始化");
        }

        #region 工作簿缓存

        /// <summary>
        /// 获取或添加工作簿缓存
        /// </summary>
        public CachedWorkbook GetOrAddWorkbook(string workbookPath, Func<Excel.Workbook> workbookFactory)
        {
            try
            {
                var cacheKey = GetWorkbookCacheKey(workbookPath);
                
                return _workbookCache.GetOrAdd(cacheKey, key =>
                {
                    var workbook = workbookFactory();
                    var cachedWorkbook = new CachedWorkbook
                    {
                        Path = workbookPath,
                        Name = workbook.Name,
                        Workbook = workbook,
                        FileSize = GetFileSize(workbookPath),
                        LastAccessTime = DateTime.Now,
                        ExpirationTime = DateTime.Now.Add(_defaultExpiration),
                        IsValid = true
                    };
                    
                    Logger.LogInfo($"工作簿已缓存: {workbook.Name} ({cachedWorkbook.FileSize:N0} bytes)");
                    return cachedWorkbook;
                });
            }
            catch (Exception ex)
            {
                Logger.LogError($"获取工作簿缓存失败: {ex.Message}", ex);
                return null;
            }
        }

        /// <summary>
        /// 获取工作簿缓存
        /// </summary>
        public CachedWorkbook GetWorkbook(string workbookPath)
        {
            try
            {
                var cacheKey = GetWorkbookCacheKey(workbookPath);
                if (_workbookCache.TryGetValue(cacheKey, out var cachedWorkbook))
                {
                    if (IsValid(cachedWorkbook))
                    {
                        cachedWorkbook.LastAccessTime = DateTime.Now;
                        return cachedWorkbook;
                    }
                    else
                    {
                        _workbookCache.TryRemove(cacheKey, out _);
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Logger.LogError($"获取工作簿缓存失败: {ex.Message}", ex);
                return null;
            }
        }

        #endregion

        #region 工作表缓存

        /// <summary>
        /// 获取或添加工作表缓存
        /// </summary>
        public CachedWorksheet GetOrAddWorksheet(string workbookPath, string sheetName, Func<Excel.Worksheet> worksheetFactory)
        {
            try
            {
                var cacheKey = GetWorksheetCacheKey(workbookPath, sheetName);
                
                return _worksheetCache.GetOrAdd(cacheKey, key =>
                {
                    var worksheet = worksheetFactory();
                    var cachedWorksheet = new CachedWorksheet
                    {
                        WorkbookPath = workbookPath,
                        SheetName = sheetName,
                        Worksheet = worksheet,
                        RowCount = GetWorksheetRowCount(worksheet),
                        ColumnCount = GetWorksheetColumnCount(worksheet),
                        LastAccessTime = DateTime.Now,
                        ExpirationTime = DateTime.Now.Add(_defaultExpiration),
                        IsValid = true
                    };
                    
                    Logger.LogInfo($"工作表已缓存: {sheetName} ({cachedWorksheet.RowCount:N0} 行, {cachedWorksheet.ColumnCount:N0} 列)");
                    return cachedWorksheet;
                });
            }
            catch (Exception ex)
            {
                Logger.LogError($"获取工作表缓存失败: {ex.Message}", ex);
                return null;
            }
        }

        /// <summary>
        /// 获取工作表缓存
        /// </summary>
        public CachedWorksheet GetWorksheet(string workbookPath, string sheetName)
        {
            try
            {
                var cacheKey = GetWorksheetCacheKey(workbookPath, sheetName);
                if (_worksheetCache.TryGetValue(cacheKey, out var cachedWorksheet))
                {
                    if (IsValid(cachedWorksheet))
                    {
                        cachedWorksheet.LastAccessTime = DateTime.Now;
                        return cachedWorksheet;
                    }
                    else
                    {
                        _worksheetCache.TryRemove(cacheKey, out _);
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Logger.LogError($"获取工作表缓存失败: {ex.Message}", ex);
                return null;
            }
        }

        #endregion

        #region 列信息缓存

        /// <summary>
        /// 获取或添加列信息缓存
        /// </summary>
        public CachedColumnInfo GetOrAddColumnInfo(string workbookPath, string sheetName, Func<List<ColumnInfo>> columnInfoFactory)
        {
            try
            {
                var cacheKey = GetColumnCacheKey(workbookPath, sheetName);
                
                return _columnCache.GetOrAdd(cacheKey, key =>
                {
                    var columns = columnInfoFactory();
                    var cachedColumns = new CachedColumnInfo
                    {
                        WorkbookPath = workbookPath,
                        SheetName = sheetName,
                        Columns = columns,
                        ColumnCount = columns.Count,
                        LastAccessTime = DateTime.Now,
                        ExpirationTime = DateTime.Now.Add(_defaultExpiration),
                        IsValid = true
                    };
                    
                    Logger.LogInfo($"列信息已缓存: {sheetName} ({cachedColumns.ColumnCount:N0} 列)");
                    return cachedColumns;
                });
            }
            catch (Exception ex)
            {
                Logger.LogError($"获取列信息缓存失败: {ex.Message}", ex);
                return null;
            }
        }

        /// <summary>
        /// 获取列信息缓存
        /// </summary>
        public CachedColumnInfo GetColumnInfo(string workbookPath, string sheetName)
        {
            try
            {
                var cacheKey = GetColumnCacheKey(workbookPath, sheetName);
                if (_columnCache.TryGetValue(cacheKey, out var cachedColumns))
                {
                    if (IsValid(cachedColumns))
                    {
                        cachedColumns.LastAccessTime = DateTime.Now;
                        return cachedColumns;
                    }
                    else
                    {
                        _columnCache.TryRemove(cacheKey, out _);
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Logger.LogError($"获取列信息缓存失败: {ex.Message}", ex);
                return null;
            }
        }

        #endregion

        #region 缓存管理

        /// <summary>
        /// 清理过期缓存
        /// </summary>
        private void CleanupExpiredCache(object state)
        {
            try
            {
                var now = DateTime.Now;
                var expiredWorkbooks = _workbookCache.Where(kvp => !IsValid(kvp.Value)).ToList();
                var expiredWorksheets = _worksheetCache.Where(kvp => !IsValid(kvp.Value)).ToList();
                var expiredColumns = _columnCache.Where(kvp => !IsValid(kvp.Value)).ToList();

                foreach (var kvp in expiredWorkbooks)
                {
                    _workbookCache.TryRemove(kvp.Key, out _);
                }

                foreach (var kvp in expiredWorksheets)
                {
                    _worksheetCache.TryRemove(kvp.Key, out _);
                }

                foreach (var kvp in expiredColumns)
                {
                    _columnCache.TryRemove(kvp.Key, out _);
                }

                if (expiredWorkbooks.Count > 0 || expiredWorksheets.Count > 0 || expiredColumns.Count > 0)
                {
                    Logger.LogInfo($"缓存清理完成: 工作簿 {expiredWorkbooks.Count}, 工作表 {expiredWorksheets.Count}, 列信息 {expiredColumns.Count}");
                }

                // 限制缓存大小
                LimitCacheSize();
            }
            catch (Exception ex)
            {
                Logger.LogError($"缓存清理失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 限制缓存大小
        /// </summary>
        private void LimitCacheSize()
        {
            try
            {
                // 限制工作簿缓存
                if (_workbookCache.Count > _maxWorkbooks)
                {
                    var toRemove = _workbookCache.OrderBy(kvp => kvp.Value.LastAccessTime)
                                                .Take(_workbookCache.Count - _maxWorkbooks)
                                                .ToList();
                    foreach (var kvp in toRemove)
                    {
                        _workbookCache.TryRemove(kvp.Key, out _);
                    }
                }

                // 限制工作表缓存
                if (_worksheetCache.Count > _maxWorksheets)
                {
                    var toRemove = _worksheetCache.OrderBy(kvp => kvp.Value.LastAccessTime)
                                                .Take(_worksheetCache.Count - _maxWorksheets)
                                                .ToList();
                    foreach (var kvp in toRemove)
                    {
                        _worksheetCache.TryRemove(kvp.Key, out _);
                    }
                }

                // 限制列信息缓存
                if (_columnCache.Count > _maxColumns)
                {
                    var toRemove = _columnCache.OrderBy(kvp => kvp.Value.LastAccessTime)
                                              .Take(_columnCache.Count - _maxColumns)
                                              .ToList();
                    foreach (var kvp in toRemove)
                    {
                        _columnCache.TryRemove(kvp.Key, out _);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"限制缓存大小失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 清除所有缓存
        /// </summary>
        public void ClearAllCache()
        {
            try
            {
                _workbookCache.Clear();
                _worksheetCache.Clear();
                _columnCache.Clear();
                
                Logger.LogInfo("所有缓存已清除");
            }
            catch (Exception ex)
            {
                Logger.LogError($"清除缓存失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 清除指定工作簿的缓存
        /// </summary>
        public void ClearWorkbookCache(string workbookPath)
        {
            try
            {
                var workbookKey = GetWorkbookCacheKey(workbookPath);
                _workbookCache.TryRemove(workbookKey, out _);

                var worksheetKeys = _worksheetCache.Keys.Where(k => k.StartsWith(workbookPath)).ToList();
                foreach (var key in worksheetKeys)
                {
                    _worksheetCache.TryRemove(key, out _);
                }

                var columnKeys = _columnCache.Keys.Where(k => k.StartsWith(workbookPath)).ToList();
                foreach (var key in columnKeys)
                {
                    _columnCache.TryRemove(key, out _);
                }

                Logger.LogInfo($"工作簿缓存已清除: {workbookPath}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"清除工作簿缓存失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 获取缓存统计信息
        /// </summary>
        public CacheStatistics GetCacheStatistics()
        {
            return new CacheStatistics
            {
                WorkbookCount = _workbookCache.Count,
                WorksheetCount = _worksheetCache.Count,
                ColumnInfoCount = _columnCache.Count,
                TotalMemoryUsage = EstimateMemoryUsage(),
                LastCleanupTime = DateTime.Now
            };
        }

        #endregion

        #region 辅助方法

        private string GetWorkbookCacheKey(string workbookPath) => $"WB_{workbookPath}";
        private string GetWorksheetCacheKey(string workbookPath, string sheetName) => $"WS_{workbookPath}_{sheetName}";
        private string GetColumnCacheKey(string workbookPath, string sheetName) => $"COL_{workbookPath}_{sheetName}";

        private bool IsValid(ICacheItem item) => item != null && item.IsValid && DateTime.Now < item.ExpirationTime;

        private long GetFileSize(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    var fileInfo = new FileInfo(filePath);
                    return fileInfo.Length;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private int GetWorksheetRowCount(Excel.Worksheet worksheet)
        {
            try
            {
                var usedRange = worksheet.UsedRange;
                return usedRange?.Rows?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private int GetWorksheetColumnCount(Excel.Worksheet worksheet)
        {
            try
            {
                var usedRange = worksheet.UsedRange;
                return usedRange?.Columns?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private long EstimateMemoryUsage()
        {
            try
            {
                long total = 0;
                total += _workbookCache.Count * 1024; // 每个工作簿约1KB
                total += _worksheetCache.Count * 512;  // 每个工作表约512B
                total += _columnCache.Count * 256;     // 每个列信息约256B
                return total;
            }
            catch
            {
                return 0;
            }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!_disposed)
            {
                _cleanupTimer?.Dispose();
                ClearAllCache();
                _disposed = true;
            }
        }

        #endregion
    }

    #region 缓存项接口和类

    /// <summary>
    /// 缓存项接口
    /// </summary>
    public interface ICacheItem
    {
        DateTime LastAccessTime { get; set; }
        DateTime ExpirationTime { get; set; }
        bool IsValid { get; set; }
    }

    /// <summary>
    /// 缓存的工作簿信息
    /// </summary>
    public class CachedWorkbook : ICacheItem
    {
        public string Path { get; set; }
        public string Name { get; set; }
        public Excel.Workbook Workbook { get; set; }
        public long FileSize { get; set; }
        public DateTime LastAccessTime { get; set; }
        public DateTime ExpirationTime { get; set; }
        public bool IsValid { get; set; }
    }

    /// <summary>
    /// 缓存的工作表信息
    /// </summary>
    public class CachedWorksheet : ICacheItem
    {
        public string WorkbookPath { get; set; }
        public string SheetName { get; set; }
        public Excel.Worksheet Worksheet { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public DateTime LastAccessTime { get; set; }
        public DateTime ExpirationTime { get; set; }
        public bool IsValid { get; set; }
    }

    /// <summary>
    /// 缓存的列信息
    /// </summary>
    public class CachedColumnInfo : ICacheItem
    {
        public string WorkbookPath { get; set; }
        public string SheetName { get; set; }
        public List<ColumnInfo> Columns { get; set; }
        public int ColumnCount { get; set; }
        public DateTime LastAccessTime { get; set; }
        public DateTime ExpirationTime { get; set; }
        public bool IsValid { get; set; }
    }

    /// <summary>
    /// 缓存统计信息
    /// </summary>
    public class CacheStatistics
    {
        public int WorkbookCount { get; set; }
        public int WorksheetCount { get; set; }
        public int ColumnInfoCount { get; set; }
        public long TotalMemoryUsage { get; set; }
        public DateTime LastCleanupTime { get; set; }
    }

    #endregion
}