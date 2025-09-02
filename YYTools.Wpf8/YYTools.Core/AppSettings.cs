using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace YYTools
{
    /// <summary>
    /// 应用设置（移植），用于WPF绑定与持久化
    /// </summary>
    public class AppSettings
    {
        private static AppSettings? _instance;
        private static readonly object _lock = new object();

        private AppSettings() { ResetToDefaults(); }

        public static AppSettings Instance
        {
            get
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new AppSettings();
                        _instance.Load();
                    }
                    return _instance;
                }
            }
        }

        public int FontSize { get; set; }
        public bool AutoScaleUI { get; set; }
        public string LogDirectory { get; set; } = string.Empty;
        public bool EnableModernUI { get; set; }
        public string Theme { get; set; } = "Default";

        public string ConcatenationDelimiter { get; set; } = "、";
        public bool RemoveDuplicateItems { get; set; } = true;
        public int MaxThreads { get; set; }
        public SortOption SortOption { get; set; }

        public bool EnableSmartColumnSelection { get; set; }
        public bool EnableColumnPreview { get; set; }
        public bool EnableColumnSearch { get; set; }

        public bool EnableColumnDataPreview { get; set; }
        public bool EnableWritePreview { get; set; }

        public int BatchSize { get; set; }
        public bool EnableProgressReporting { get; set; }
        public int MaxRowsForPreview { get; set; }
        public int PreviewParseRows { get; set; }

        public bool EnableCaching { get; set; }
        public int CacheExpirationMinutes { get; set; }
        public int MaxCachedWorkbooks { get; set; }
        public int MaxCachedWorksheets { get; set; }
        public int MaxCachedColumns { get; set; }

        public bool EnableSmartMatching { get; set; }
        public bool EnableExactMatchPriority { get; set; }
        public double MinMatchScore { get; set; }

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
                if (!File.Exists(ConfigPath)) { Save(); return; }
                var dict = File.ReadAllLines(ConfigPath)
                    .Where(l => !string.IsNullOrWhiteSpace(l) && !l.StartsWith("#") && l.Contains('='))
                    .Select(l => l.Split(new[] { '=' }, 2))
                    .ToDictionary(p => p[0].Trim(), p => p[1].Trim(), StringComparer.OrdinalIgnoreCase);

                GetValue(dict, "FontSize", v => FontSize = int.Parse(v));
                GetValue(dict, "AutoScaleUI", v => AutoScaleUI = bool.Parse(v));
                GetValue(dict, "LogDirectory", v => LogDirectory = v);
                GetValue(dict, "EnableModernUI", v => EnableModernUI = bool.Parse(v));
                GetValue(dict, "Theme", v => Theme = v);

                GetValue(dict, "ConcatenationDelimiter", v => ConcatenationDelimiter = v);
                GetValue(dict, "RemoveDuplicateItems", v => RemoveDuplicateItems = bool.Parse(v));
                GetValue(dict, "MaxThreads", v => MaxThreads = int.Parse(v));
                GetValue(dict, "SortOption", v => { if (Enum.TryParse<SortOption>(v, true, out var s)) SortOption = s; });

                GetValue(dict, "EnableSmartColumnSelection", v => EnableSmartColumnSelection = bool.Parse(v));
                GetValue(dict, "EnableColumnPreview", v => EnableColumnPreview = bool.Parse(v));
                GetValue(dict, "EnableColumnSearch", v => EnableColumnSearch = bool.Parse(v));

                GetValue(dict, "BatchSize", v => BatchSize = int.Parse(v));
                GetValue(dict, "EnableProgressReporting", v => EnableProgressReporting = bool.Parse(v));
                GetValue(dict, "MaxRowsForPreview", v => MaxRowsForPreview = int.Parse(v));
                GetValue(dict, "PreviewParseRows", v => PreviewParseRows = int.Parse(v));

                GetValue(dict, "EnableCaching", v => EnableCaching = bool.Parse(v));
                GetValue(dict, "CacheExpirationMinutes", v => CacheExpirationMinutes = int.Parse(v));
                GetValue(dict, "MaxCachedWorkbooks", v => MaxCachedWorkbooks = int.Parse(v));
                GetValue(dict, "MaxCachedWorksheets", v => MaxCachedWorksheets = int.Parse(v));
                GetValue(dict, "MaxCachedColumns", v => MaxCachedColumns = int.Parse(v));

                GetValue(dict, "EnableSmartMatching", v => EnableSmartMatching = bool.Parse(v));
                GetValue(dict, "EnableExactMatchPriority", v => EnableExactMatchPriority = bool.Parse(v));
                GetValue(dict, "MinMatchScore", v => MinMatchScore = double.Parse(v));

                GetValue(dict, "EnableAsyncProcessing", v => EnableAsyncProcessing = bool.Parse(v));
                GetValue(dict, "AsyncTaskTimeoutSeconds", v => AsyncTaskTimeoutSeconds = int.Parse(v));
                GetValue(dict, "EnableBackgroundTasks", v => EnableBackgroundTasks = bool.Parse(v));

                GetValue(dict, "EnableColumnDataPreview", v => EnableColumnDataPreview = bool.Parse(v));
                GetValue(dict, "EnableWritePreview", v => EnableWritePreview = bool.Parse(v));
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
                try { assign(dict[key]); }
                catch (Exception ex) { Logger.LogWarning($"设置值转换失败: {key} = {dict[key]}, 错误: {ex.Message}"); }
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
                    $"SortOption={SortOption}",
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
                    $"PreviewParseRows={PreviewParseRows}",
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
                    "",
                    "# 性能优化设置",
                    $"EnableColumnDataPreview={EnableColumnDataPreview}",
                    $"EnableWritePreview={EnableWritePreview}",
                    $"BatchSize={BatchSize}",
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
            SortOption = SortOption.None;

            EnableSmartColumnSelection = true;
            EnableColumnPreview = true;
            EnableColumnSearch = true;

            BatchSize = Constants.DefaultBatchSize;
            EnableProgressReporting = true;
            MaxRowsForPreview = Constants.DefaultMaxPreviewRows;
            PreviewParseRows = Constants.DefaultPreviewParseRows;

            EnableCaching = true;
            CacheExpirationMinutes = Constants.DefaultCacheExpirationMinutes;
            MaxCachedWorkbooks = Constants.MaxCachedWorkbooks;
            MaxCachedWorksheets = Constants.MaxCachedWorksheets;
            MaxCachedColumns = Constants.MaxCachedColumns;

            EnableSmartMatching = true;
            EnableExactMatchPriority = true;
            MinMatchScore = 0.5;

            EnableAsyncProcessing = true;
            AsyncTaskTimeoutSeconds = 300;
            EnableBackgroundTasks = true;

            EnableColumnDataPreview = true;
            EnableWritePreview = true;
            BatchSize = Constants.DefaultBatchSize;
        }
    }
}

