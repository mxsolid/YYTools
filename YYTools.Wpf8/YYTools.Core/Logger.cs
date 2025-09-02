using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace YYTools
{
    /// <summary>
    /// 跨项目复用的日志记录器（移植并精简）
    /// </summary>
    public static class Logger
    {
        private static string _logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");
        private static readonly object _lockObject = new object();
        private static readonly Queue<LogEntry> _queue = new Queue<LogEntry>();
        private static Timer? _flushTimer;
        private static bool _initialized;

        static Logger()
        {
            try
            {
                Directory.CreateDirectory(_logPath);
                _flushTimer = new Timer(_ => Flush(), null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));
                _initialized = true;
                Log("日志系统已初始化", LogLevel.Info);
            }
            catch
            {
                _initialized = false;
            }
        }

        public static void LogInfo(string message) => Log(message, LogLevel.Info);
        public static void LogWarning(string message) => Log(message, LogLevel.Warning);
        public static void LogDebug(string message) => Log(message, LogLevel.Debug);
        public static void LogError(string message, Exception? ex = null)
        {
            var msg = ex != null ? $"{message}\n详情: {ex}" : message;
            Log(msg, LogLevel.Error);
        }

        public static void Log(string message, LogLevel level)
        {
            if (!_initialized)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}");
                return;
            }

            lock (_lockObject)
            {
                _queue.Enqueue(new LogEntry
                {
                    Timestamp = DateTime.Now,
                    Level = level,
                    Message = message,
                    ThreadId = Thread.CurrentThread.ManagedThreadId
                });

                if (_queue.Count > 200)
                {
                    Flush();
                }
            }
        }

        public static void ForceFlush() => Flush();

        private static void Flush()
        {
            List<LogEntry> toWrite;
            lock (_lockObject)
            {
                if (_queue.Count == 0) return;
                toWrite = _queue.ToList();
                _queue.Clear();
            }

            var sb = new StringBuilder();
            foreach (var e in toWrite)
            {
                sb.AppendLine($"[{e.Timestamp:yyyy-MM-dd HH:mm:ss.fff}] [{e.Level}] [线程:{e.ThreadId}] {e.Message}");
            }

            try
            {
                var file = Path.Combine(_logPath, $"YYTools_{DateTime.Now:yyyy-MM-dd}.log");
                File.AppendAllText(file, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
                // 忽略写入失败
            }
        }
    }

    public class LogEntry
    {
        public DateTime Timestamp { get; set; }
        public LogLevel Level { get; set; }
        public string Message { get; set; } = string.Empty;
        public int ThreadId { get; set; }
    }
}

