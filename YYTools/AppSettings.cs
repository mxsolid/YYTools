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

        // 运单匹配工具的独立设置
        public string ConcatenationDelimiter { get; set; }
        public bool RemoveDuplicateItems { get; set; }
        public int MaxThreads { get; set; }

        private string ConfigPath
        {
            get
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                return Path.Combine(folder, "settings.ini");
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
                GetValue(settingsDict, "ConcatenationDelimiter", v => ConcatenationDelimiter = v);
                GetValue(settingsDict, "RemoveDuplicateItems", v => RemoveDuplicateItems = bool.Parse(v));
                GetValue(settingsDict, "MaxThreads", v => MaxThreads = int.Parse(v));

            }
            catch (Exception)
            {
                ResetToDefaults();
                Save();
            }
        }

        private void GetValue(Dictionary<string, string> dict, string key, Action<string> assign)
        {
            if (dict.ContainsKey(key))
            {
                assign(dict[key]);
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
                    "",
                    "# 运单匹配工具设置",
                    $"ConcatenationDelimiter={ConcatenationDelimiter}",
                    $"RemoveDuplicateItems={RemoveDuplicateItems}",
                    $"MaxThreads={MaxThreads}",
                };
                File.WriteAllLines(ConfigPath, lines, System.Text.Encoding.UTF8);
            }
            catch { }
        }

        public void ResetToDefaults()
        {
            FontSize = 9;
            AutoScaleUI = true;
            LogDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "YYTools", "Logs");
            MaxThreads = Environment.ProcessorCount;
            ConcatenationDelimiter = "、";
            RemoveDuplicateItems = true;
        }
    }
}