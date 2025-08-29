using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace YYTools
{
    /// <summary>
    /// 增强的日志记录系统
    /// </summary>
    public static class Logger
    {
        private static readonly string LogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");
        private static readonly string PerformanceLogPath = Path.Combine(LogPath, "Performance");
        private static readonly string UserActionLogPath = Path.Combine(LogPath, "UserActions");
        private static readonly string ErrorLogPath = Path.Combine(LogPath, "Errors");
        
        private static readonly object lockObject = new object();
        private static readonly Queue<LogEntry> logQueue = new Queue<LogEntry>();
        private static readonly Timer flushTimer;
        private static bool isInitialized = false;

        static Logger()
        {
            try
            {
                InitializeLogDirectories();
                flushTimer = new Timer(FlushLogQueue, null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));
                isInitialized = true;
                
                // 立即记录初始化日志
                Log("日志系统已初始化", LogLevel.Info);
            }
            catch (Exception ex)
            {
                // 如果初始化失败，尝试写入系统临时目录
                try
                {
                    LogPath = Path.Combine(Path.GetTempPath(), "YYTools", "Logs");
                    InitializeLogDirectories();
                    isInitialized = true;
                    Log("日志系统已初始化（使用临时目录）", LogLevel.Info);
                }
                catch
                {
                    // 如果还是失败，就使用当前目录
                    LogPath = "Logs";
                    try
                    {
                        if (!Directory.Exists(LogPath))
                            Directory.CreateDirectory(LogPath);
                        isInitialized = true;
                        Log("日志系统已初始化（使用当前目录）", LogLevel.Info);
                    }
                    catch
                    {
                        isInitialized = false;
                        // 最后尝试直接写入控制台
                        Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 日志系统初始化失败: {ex.Message}");
                    }
                }
            }
        }

        #region 基础日志方法

        /// <summary>
        /// 记录信息日志
        /// </summary>
        public static void LogInfo(string message)
        {
            Log(message, LogLevel.Info);
        }

        /// <summary>
        /// 记录警告日志
        /// </summary>
        public static void LogWarning(string message)
        {
            Log(message, LogLevel.Warning);
        }

        /// <summary>
        /// 记录错误日志
        /// </summary>
        public static void LogError(string message, Exception ex = null)
        {
            string fullMessage = ex != null ? $"{message}\n详情: {ex}" : message;
            Log(fullMessage, LogLevel.Error);
            
            // 错误日志同时写入错误专用文件
            if (ex != null)
            {
                WriteErrorLog(message, ex);
            }
        }

        /// <summary>
        /// 记录调试日志
        /// </summary>
        public static void LogDebug(string message)
        {
            Log(message, LogLevel.Debug);
        }

        /// <summary>
        /// 记录性能日志
        /// </summary>
        public static void LogPerformance(string operation, TimeSpan duration, string details = "")
        {
            var performanceLog = new PerformanceLogEntry
            {
                Timestamp = DateTime.Now,
                Operation = operation,
                Duration = duration,
                Details = details
            };
            
            WritePerformanceLog(performanceLog);
        }

        #endregion

        #region 用户操作日志

        /// <summary>
        /// 记录用户操作
        /// </summary>
        public static void LogUserAction(string action, string details = "", string result = "")
        {
            var userActionLog = new UserActionLogEntry
            {
                Timestamp = DateTime.Now,
                Action = action,
                Details = details,
                Result = result,
                UserName = Environment.UserName,
                MachineName = Environment.MachineName
            };
            
            WriteUserActionLog(userActionLog);
            
            // 同时记录到主日志
            Log($"[用户操作] {action} - {details} - {result}", LogLevel.Info);
        }

        /// <summary>
        /// 记录文件操作
        /// </summary>
        public static void LogFileOperation(string operation, string filePath, long fileSize = 0, string result = "")
        {
            var details = $"文件: {filePath}";
            if (fileSize > 0)
            {
                details += $", 大小: {FormatFileSize(fileSize)}";
            }
            
            LogUserAction($"文件{operation}", details, result);
        }

        /// <summary>
        /// 记录Excel操作
        /// </summary>
        public static void LogExcelOperation(string operation, string workbookName, string sheetName = "", int rowCount = 0, int columnCount = 0)
        {
            var details = $"工作簿: {workbookName}";
            if (!string.IsNullOrEmpty(sheetName))
            {
                details += $", 工作表: {sheetName}";
            }
            if (rowCount > 0)
            {
                details += $", 行数: {rowCount:N0}";
            }
            if (columnCount > 0)
            {
                details += $", 列数: {columnCount:N0}";
            }
            
            LogUserAction($"Excel{operation}", details, "成功");
        }

        /// <summary>
        /// 记录配置变更
        /// </summary>
        public static void LogConfigurationChange(string settingName, string oldValue, string newValue)
        {
            var details = $"设置: {settingName}, 原值: {oldValue}, 新值: {newValue}";
            LogUserAction("配置变更", details, "已保存");
        }

        #endregion

        #region 核心日志方法

        /// <summary>
        /// 记录日志到队列
        /// </summary>
        public static void Log(string message, LogLevel level)
        {
            if (!isInitialized)
            {
                // 如果日志系统未初始化，直接写入控制台
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}");
                return;
            }

            var logEntry = new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = level,
                Message = message,
                ThreadId = Thread.CurrentThread.ManagedThreadId
            };

            lock (lockObject)
            {
                logQueue.Enqueue(logEntry);
                
                // 如果队列太长，立即刷新
                if (logQueue.Count > 100)
                {
                    FlushLogQueue(null);
                }
            }
        }

        /// <summary>
        /// 刷新日志队列
        /// </summary>
        private static void FlushLogQueue(object state)
        {
            if (logQueue.Count == 0) return;

            var entriesToWrite = new List<LogEntry>();
            
            lock (lockObject)
            {
                while (logQueue.Count > 0)
                {
                    entriesToWrite.Add(logQueue.Dequeue());
                }
            }

            if (entriesToWrite.Count > 0)
            {
                WriteLogEntries(entriesToWrite);
            }
        }

        /// <summary>
        /// 写入日志条目
        /// </summary>
        private static void WriteLogEntries(List<LogEntry> entries)
        {
            try
            {
                string logFile = Path.Combine(LogPath, $"YYTools_{DateTime.Now:yyyy-MM-dd}.log");
                var logBuilder = new StringBuilder();

                foreach (var entry in entries)
                {
                    logBuilder.AppendLine(FormatLogEntry(entry));
                }

                File.AppendAllText(logFile, logBuilder.ToString(), Encoding.UTF8);
            }
            catch (Exception _)
            {
                // 如果写入失败，尝试写入备用位置
                try
                {
                    string backupLogFile = Path.Combine(Path.GetTempPath(), $"YYTools_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                    var backupBuilder = new StringBuilder();
                    foreach (var entry in entries)
                    {
                        backupBuilder.AppendLine(FormatLogEntry(entry));
                    }
                    File.WriteAllText(backupLogFile, backupBuilder.ToString(), Encoding.UTF8);
                }
                catch
                {
                    // 如果备用位置也失败，就忽略
                }
            }
        }

        #endregion

        #region 专用日志写入

        /// <summary>
        /// 写入性能日志
        /// </summary>
        private static void WritePerformanceLog(PerformanceLogEntry entry)
        {
            try
            {
                string performanceFile = Path.Combine(PerformanceLogPath, $"Performance_{DateTime.Now:yyyy-MM-dd}.log");
                string logLine = $"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff}] {entry.Operation} | 耗时: {entry.Duration.TotalMilliseconds:F2}ms | {entry.Details}";
                File.AppendAllText(performanceFile, logLine + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // 忽略性能日志写入失败
            }
        }

        /// <summary>
        /// 写入用户操作日志
        /// </summary>
        private static void WriteUserActionLog(UserActionLogEntry entry)
        {
            try
            {
                string userActionFile = Path.Combine(UserActionLogPath, $"UserActions_{DateTime.Now:yyyy-MM-dd}.log");
                string logLine = $"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss}] 用户: {entry.UserName} | 机器: {entry.MachineName} | 操作: {entry.Action} | 详情: {entry.Details} | 结果: {entry.Result}";
                File.AppendAllText(userActionFile, logLine + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // 忽略用户操作日志写入失败
            }
        }

        /// <summary>
        /// 写入错误日志
        /// </summary>
        private static void WriteErrorLog(string message, Exception ex)
        {
            try
            {
                string errorFile = Path.Combine(ErrorLogPath, $"Errors_{DateTime.Now:yyyy-MM-dd}.log");
                var errorBuilder = new StringBuilder();
                errorBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 错误: {message}");
                errorBuilder.AppendLine($"异常类型: {ex.GetType().Name}");
                errorBuilder.AppendLine($"异常消息: {ex.Message}");
                errorBuilder.AppendLine($"堆栈跟踪: {ex.StackTrace}");
                errorBuilder.AppendLine(new string('-', 80));
                
                File.AppendAllText(errorFile, errorBuilder.ToString(), Encoding.UTF8);
            }
            catch
            {
                // 忽略错误日志写入失败
            }
        }

        #endregion

        #region 日志管理

        /// <summary>
        /// 清理旧日志文件
        /// </summary>
        public static void CleanupOldLogs()
        {
            try
            {
                CleanupLogDirectory(LogPath, Constants.MaxLogFiles, Constants.MaxLogFileSizeMB);
                CleanupLogDirectory(PerformanceLogPath, Constants.MaxLogFiles, Constants.MaxLogFileSizeMB);
                CleanupLogDirectory(UserActionLogPath, Constants.MaxLogFiles, Constants.MaxLogFileSizeMB);
                CleanupLogDirectory(ErrorLogPath, Constants.MaxLogFiles, Constants.MaxLogFileSizeMB);
            }
            catch (Exception ex)
            {
                // 记录清理失败，但不抛出异常
                Log($"清理旧日志文件失败: {ex.Message}", LogLevel.Warning);
            }
        }

        /// <summary>
        /// 清理指定目录的旧日志文件
        /// </summary>
        private static void CleanupLogDirectory(string directory, int maxFiles, int maxFileSizeMB)
        {
            if (!Directory.Exists(directory)) return;

            try
            {
                var logFiles = Directory.GetFiles(directory, "*.log")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();

                // 删除超过最大文件数的旧文件
                if (logFiles.Count > maxFiles)
                {
                    var filesToDelete = logFiles.Skip(maxFiles);
                    foreach (var file in filesToDelete)
                    {
                        try
                        {
                            file.Delete();
                        }
                        catch
                        {
                            // 忽略单个文件删除失败
                        }
                    }
                }

                // 删除超过最大文件大小的文件
                var largeFiles = logFiles.Where(f => f.Length > maxFileSizeMB * 1024 * 1024);
                foreach (var file in largeFiles)
                {
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                        // 忽略单个文件删除失败
                    }
                }
            }
            catch
            {
                // 忽略目录清理失败
            }
        }

        /// <summary>
        /// 获取日志统计信息
        /// </summary>
        public static LogStatistics GetLogStatistics()
        {
            try
            {
                var stats = new LogStatistics();
                
                if (Directory.Exists(LogPath))
                {
                    var logFiles = Directory.GetFiles(LogPath, "*.log");
                    stats.TotalLogFiles = logFiles.Length;
                    stats.TotalLogSize = logFiles.Sum(f => new FileInfo(f).Length);
                }

                if (Directory.Exists(PerformanceLogPath))
                {
                    var perfFiles = Directory.GetFiles(PerformanceLogPath, "*.log");
                    stats.PerformanceLogFiles = perfFiles.Length;
                }

                if (Directory.Exists(UserActionLogPath))
                {
                    var actionFiles = Directory.GetFiles(UserActionLogPath, "*.log");
                    stats.UserActionLogFiles = actionFiles.Length;
                }

                if (Directory.Exists(ErrorLogPath))
                {
                    var errorFiles = Directory.GetFiles(ErrorLogPath, "*.log");
                    stats.ErrorLogFiles = errorFiles.Length;
                }

                return stats;
            }
            catch
            {
                return new LogStatistics();
            }
        }

        /// <summary>
        /// 强制刷新日志队列
        /// </summary>
        public static void ForceFlush()
        {
            FlushLogQueue(null);
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 初始化日志目录
        /// </summary>
        private static void InitializeLogDirectories()
        {
            var directories = new[] { LogPath, PerformanceLogPath, UserActionLogPath, ErrorLogPath };
            
            foreach (var directory in directories)
            {
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
        }

        /// <summary>
        /// 格式化日志条目
        /// </summary>
        private static string FormatLogEntry(LogEntry entry)
        {
            return $"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff}] [{entry.Level}] [线程:{entry.ThreadId}] {entry.Message}";
        }

        /// <summary>
        /// 格式化文件大小
        /// </summary>
        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double size = bytes;
            
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }

            return $"{size:0.##} {sizes[order]}";
        }

        #endregion
    }

    #region 日志相关类

    /// <summary>
    /// 日志条目
    /// </summary>
    public class LogEntry
    {
        public DateTime Timestamp { get; set; }
        public LogLevel Level { get; set; }
        public string Message { get; set; }
        public int ThreadId { get; set; }
    }

    /// <summary>
    /// 性能日志条目
    /// </summary>
    public class PerformanceLogEntry
    {
        public DateTime Timestamp { get; set; }
        public string Operation { get; set; }
        public TimeSpan Duration { get; set; }
        public string Details { get; set; }
    }

    /// <summary>
    /// 用户操作日志条目
    /// </summary>
    public class UserActionLogEntry
    {
        public DateTime Timestamp { get; set; }
        public string Action { get; set; }
        public string Details { get; set; }
        public string Result { get; set; }
        public string UserName { get; set; }
        public string MachineName { get; set; }
    }

    /// <summary>
    /// 日志统计信息
    /// </summary>
    public class LogStatistics
    {
        public int TotalLogFiles { get; set; }
        public long TotalLogSize { get; set; }
        public int PerformanceLogFiles { get; set; }
        public int UserActionLogFiles { get; set; }
        public int ErrorLogFiles { get; set; }
    }

    #endregion
}