using System;
using System.IO;

namespace YYTools
{
    public static class Logger
    {
        private static readonly string LogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");
        private static readonly object lockObject = new object();

        static Logger()
        {
            try
            {
                if (!Directory.Exists(LogPath))
                {
                    Directory.CreateDirectory(LogPath);
                }
            }
            catch { }
        }

        public static void Log(string message, LogLevel level)
        {
            lock (lockObject)
            {
                try
                {
                    string logFile = Path.Combine(LogPath, $"YYTools_{DateTime.Now:yyyy-MM-dd}.log");
                    string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                    File.AppendAllText(logFile, logEntry + Environment.NewLine, System.Text.Encoding.UTF8);
                }
                catch { }
            }
        }

        public static void LogInfo(string message) => Log(message, LogLevel.Info);
        public static void LogWarning(string message) => Log(message, LogLevel.Warning);
        public static void LogError(string message, Exception ex = null)
        {
            string fullMessage = ex != null ? $"{message}\n详情: {ex}" : message;
            Log(fullMessage, LogLevel.Error);
        }
        
        public static void LogUserAction(string action)
        {
            Log($"[UI Action] {action}", LogLevel.Info);
        }
    }
}