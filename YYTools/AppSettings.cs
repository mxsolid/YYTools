// --- 文件 2: AppSettings.cs ---
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace YYTools
{
    public class AppSettings
    {
        private static AppSettings instance;
        private static readonly object lockObject = new object();

        private AppSettings() { ResetToDefaults(); }

        public static AppSettings Instance
        {
            get
            {
                lock (lockObject)
                {
                    if (instance == null)
                    {
                        instance = new AppSettings();
                        instance.Load();
                    }
                    return instance;
                }
            }
        }

        // 通用设置
        public int FontSize { get; set; }
        public bool AutoScaleUI { get; set; }
        public string LogDirectory { get; set; }
        public bool EnableModernUI { get; set; }
        public string Theme { get; set; }

        // 运单匹配工具的独立设置
        public string ConcatenationDelimiter { get; set; }
        public bool RemoveDuplicateItems { get; set; }
        public int MaxThreads { get; set; }
        
        // 智能列选择设置
        public bool EnableSmartColumnSelection { get; set; }
        public bool EnableColumnPreview { get; set; }
        public bool EnableColumnSearch { get; set; }
        
        // 性能优化设置
        public int BatchSize { get; set; }
        public bool EnableProgressReporting { get; set; }
        public int MaxRowsForPreview { get; set; }
        
        // 缓存设置
        public bool EnableCaching { get; set; }
        public int CacheExpirationMinutes { get; set; }
        public int MaxCachedWorkbooks { get; set; }
        public int MaxCachedWorksheets { get; set; }
        public int MaxCachedColumns { get; set; }
        
        // 智能匹配设置
        public bool EnableSmartMatching { get; set; }
        public bool EnableExactMatchPriority { get; set; }
        public double MinMatchScore { get; set; }
        
        // 异步处理设置
        public bool EnableAsyncProcessing { get; set; }
        public int AsyncTaskTimeoutSeconds { get; set; }
        public bool EnableBackgroundTasks { get; set; }

        private string ConfigPath
        {
            get
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                return Path.Combine(folder, Constants.ConfigFileName);
            }
        }

        public void Load()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    Save();
                    return;
                }

                var settingsDict = File.ReadAllLines(ConfigPath)
                    .Where(line => !string.IsNullOrWhiteSpace(line) && !line.StartsWith("#") && line.Contains("="))
                    .Select(line => line.Split(new[] { '=' }, 2))
                    .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim(), StringComparer.OrdinalIgnoreCase);

                // 安全地加载每个设置
                GetValue(settingsDict, "FontSize", v => FontSize = int.Parse(v));
                GetValue(settingsDict, "AutoScaleUI", v => AutoScaleUI = bool.Parse(v));
                GetValue(settingsDict, "LogDirectory", v => LogDirectory = v);
                GetValue(settingsDict, "EnableModernUI", v => EnableModernUI = bool.Parse(v));
                GetValue(settingsDict, "Theme", v => Theme = v);
                
                GetValue(settingsDict, "ConcatenationDelimiter", v => ConcatenationDelimiter = v);
                GetValue(settingsDict, "RemoveDuplicateItems", v => RemoveDuplicateItems = bool.Parse(v));
                GetValue(settingsDict, "MaxThreads", v => MaxThreads = int.Parse(v));
                
                // 智能列选择设置
                GetValue(settingsDict, "EnableSmartColumnSelection", v => EnableSmartColumnSelection = bool.Parse(v));
                GetValue(settingsDict, "EnableColumnPreview", v => EnableColumnPreview = bool.Parse(v));
                GetValue(settingsDict, "EnableColumnSearch", v => EnableColumnSearch = bool.Parse(v));
                
                // 性能优化设置
                GetValue(settingsDict, "BatchSize", v => BatchSize = int.Parse(v));
                GetValue(settingsDict, "EnableProgressReporting", v => EnableProgressReporting = bool.Parse(v));
                GetValue(settingsDict, "MaxRowsForPreview", v => MaxRowsForPreview = int.Parse(v));
                
                // 缓存设置
                GetValue(settingsDict, "EnableCaching", v => EnableCaching = bool.Parse(v));
                GetValue(settingsDict, "CacheExpirationMinutes", v => CacheExpirationMinutes = int.Parse(v));
                GetValue(settingsDict, "MaxCachedWorkbooks", v => MaxCachedWorkbooks = int.Parse(v));
                GetValue(settingsDict, "MaxCachedWorksheets", v => MaxCachedWorksheets = int.Parse(v));
                GetValue(settingsDict, "MaxCachedColumns", v => MaxCachedColumns = int.Parse(v));
                
                // 智能匹配设置
                GetValue(settingsDict, "EnableSmartMatching", v => EnableSmartMatching = bool.Parse(v));
                GetValue(settingsDict, "EnableExactMatchPriority", v => EnableExactMatchPriority = bool.Parse(v));
                GetValue(settingsDict, "MinMatchScore", v => MinMatchScore = double.Parse(v));
                
                // 异步处理设置
                GetValue(settingsDict, "EnableAsyncProcessing", v => EnableAsyncProcessing = bool.Parse(v));
                GetValue(settingsDict, "AsyncTaskTimeoutSeconds", v => AsyncTaskTimeoutSeconds = int.Parse(v));
                GetValue(settingsDict, "EnableBackgroundTasks", v => EnableBackgroundTasks = bool.Parse(v));

            }
            catch (Exception ex)
            {
                Logger.LogError("加载应用程序设置失败", ex);
                ResetToDefaults();
                Save();
            }
        }

        private void GetValue(Dictionary<string, string> dict, string key, Action<string> assign)
        {
            if (dict.ContainsKey(key))
            {
                try
                {
                    assign(dict[key]);
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"设置值转换失败: {key} = {dict[key]}, 错误: {ex.Message}");
                }
            }
        }

        public void Save()
        {
            try
            {
                var lines = new List<string>
                {
                    "# YY工具通用设置",
                    $"FontSize={FontSize}",
                    $"AutoScaleUI={AutoScaleUI}",
                    $"LogDirectory={LogDirectory}",
                    $"EnableModernUI={EnableModernUI}",
                    $"Theme={Theme}",
                    "",
                    "# 运单匹配工具设置",
                    $"ConcatenationDelimiter={ConcatenationDelimiter}",
                    $"RemoveDuplicateItems={RemoveDuplicateItems}",
                    $"MaxThreads={MaxThreads}",
                    "",
                    "# 智能列选择设置",
                    $"EnableSmartColumnSelection={EnableSmartColumnSelection}",
                    $"EnableColumnPreview={EnableColumnPreview}",
                    $"EnableColumnSearch={EnableColumnSearch}",
                    "",
                    "# 性能优化设置",
                    $"BatchSize={BatchSize}",
                    $"EnableProgressReporting={EnableProgressReporting}",
                    $"MaxRowsForPreview={MaxRowsForPreview}",
                    "",
                    "# 缓存设置",
                    $"EnableCaching={EnableCaching}",
                    $"CacheExpirationMinutes={CacheExpirationMinutes}",
                    $"MaxCachedWorkbooks={MaxCachedWorkbooks}",
                    $"MaxCachedWorksheets={MaxCachedWorksheets}",
                    $"MaxCachedColumns={MaxCachedColumns}",
                    "",
                    "# 智能匹配设置",
                    $"EnableSmartMatching={EnableSmartMatching}",
                    $"EnableExactMatchPriority={EnableExactMatchPriority}",
                    $"MinMatchScore={MinMatchScore}",
                    "",
                    "# 异步处理设置",
                    $"EnableAsyncProcessing={EnableAsyncProcessing}",
                    $"AsyncTaskTimeoutSeconds={AsyncTaskTimeoutSeconds}",
                    $"EnableBackgroundTasks={EnableBackgroundTasks}",
                };
                File.WriteAllLines(ConfigPath, lines, System.Text.Encoding.UTF8);
                
                Logger.LogInfo("应用程序设置已保存");
            }
            catch (Exception ex)
            {
                Logger.LogError("保存应用程序设置失败", ex);
            }
        }

        public void ResetToDefaults()
        {
            FontSize = Constants.DefaultFontSize;
            AutoScaleUI = true;
            LogDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");
            EnableModernUI = true;
            Theme = "Default";
            
            MaxThreads = Environment.ProcessorCount;
            ConcatenationDelimiter = "、";
            RemoveDuplicateItems = true;
            
            // 智能列选择默认值
            EnableSmartColumnSelection = true;
            EnableColumnPreview = true;
            EnableColumnSearch = true;
            
            // 性能优化默认值
            BatchSize = Constants.DefaultBatchSize;
            EnableProgressReporting = true;
            MaxRowsForPreview = Constants.DefaultMaxPreviewRows;
            
            // 缓存默认值
            EnableCaching = true;
            CacheExpirationMinutes = Constants.DefaultCacheExpirationMinutes;
            MaxCachedWorkbooks = Constants.MaxCachedWorkbooks;
            MaxCachedWorksheets = Constants.MaxCachedWorksheets;
            MaxCachedColumns = Constants.MaxCachedColumns;
            
            // 智能匹配默认值
            EnableSmartMatching = true;
            EnableExactMatchPriority = true;
            MinMatchScore = 0.5;
            
            // 异步处理默认值
            EnableAsyncProcessing = true;
            AsyncTaskTimeoutSeconds = 300; // 5分钟
            EnableBackgroundTasks = true;
        }
        
        /// <summary>
        /// 获取分隔符选项
        /// </summary>
        public string[] GetDelimiterOptions()
        {
            return Constants.DelimiterOptions;
        }
        
        /// <summary>
        /// 获取排序选项
        /// </summary>
        public string[] GetSortOptions()
        {
            return Constants.SortOptions;
        }
        
        /// <summary>
        /// 验证设置有效性
        /// </summary>
        public List<string> ValidateSettings()
        {
            var errors = new List<string>();
            
            try
            {
                if (FontSize < Constants.MinFontSize || FontSize > Constants.MaxFontSize)
                    errors.Add($"字体大小必须在 {Constants.MinFontSize}-{Constants.MaxFontSize} 之间");
                
                if (MaxThreads < 1 || MaxThreads > 32)
                    errors.Add("最大线程数必须在 1-32 之间");
                
                if (BatchSize < 100 || BatchSize > 10000)
                    errors.Add("批处理大小必须在 100-10000 之间");
                
                if (MaxRowsForPreview < 10 || MaxRowsForPreview > 1000)
                    errors.Add("预览最大行数必须在 10-1000 之间");
                
                if (CacheExpirationMinutes < 1 || CacheExpirationMinutes > 1440)
                    errors.Add("缓存过期时间必须在 1-1440 分钟之间");
                
                if (MinMatchScore < 0.0 || MinMatchScore > 1.0)
                    errors.Add("最小匹配分数必须在 0.0-1.0 之间");
                
                if (AsyncTaskTimeoutSeconds < 30 || AsyncTaskTimeoutSeconds > 3600)
                    errors.Add("异步任务超时时间必须在 30-3600 秒之间");
            }
            catch (Exception ex)
            {
                errors.Add($"设置验证失败: {ex.Message}");
            }
            
            return errors;
        }
    }
}