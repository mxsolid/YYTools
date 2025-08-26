using System;
using System.Collections.Generic;
using System.IO;

namespace YYTools
{
    /// <summary>
    /// 应用程序设置类 - WPS优先
    /// </summary>
    public class AppSettings
    {
        private static AppSettings instance;
        private static readonly object lockObject = new object();

        private AppSettings()
        {
            // 初始化默认值
            ResetToDefaults();
        }

        public static AppSettings Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (lockObject)
                    {
                        if (instance == null)
                        {
                            instance = new AppSettings();
                            instance.Load();
                        }
                    }
                }
                return instance;
            }
        }

        // 界面设置
        public int FontSize { get; set; }
        public bool AutoScaleUI { get; set; }

        // 性能模式
        public PerformanceMode PerformanceMode { get; set; }

        // WPS优先设置
        public bool WPSPriority { get; set; }
        public bool EnableDebugLog { get; set; }

        // 默认列设置
        public string DefaultShippingTrackColumn { get; set; }
        public string DefaultShippingProductColumn { get; set; }
        public string DefaultShippingNameColumn { get; set; }
        public string DefaultBillTrackColumn { get; set; }
        public string DefaultBillProductColumn { get; set; }
        public string DefaultBillNameColumn { get; set; }

        // 高级设置
        public int ProgressUpdateFrequency { get; set; }
        public string LogDirectory { get; set; }

        private string ConfigPath
        {
            get
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "YYTools");

                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                return Path.Combine(folder, "settings.ini");
            }
        }

        public void Load()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    Save(); // Save default settings if file doesn't exist
                    return;
                }

                string[] lines = File.ReadAllLines(ConfigPath, System.Text.Encoding.UTF8);
                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;

                    string[] parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length != 2) continue;

                    string key = parts[0].Trim();
                    string value = parts[1].Trim();

                    switch (key)
                    {
                        case "FontSize":
                            if (int.TryParse(value, out int fontSize)) FontSize = fontSize > 0 ? fontSize : 9;
                            break;
                        case "AutoScaleUI":
                            if (bool.TryParse(value, out bool autoScale)) AutoScaleUI = autoScale;
                            break;
                        case "PerformanceMode":
                            try { PerformanceMode = (PerformanceMode)Enum.Parse(typeof(PerformanceMode), value); }
                            catch { PerformanceMode = PerformanceMode.UltraFast; }
                            break;
                        case "WPSPriority":
                            if (bool.TryParse(value, out bool wpsPriority)) WPSPriority = wpsPriority;
                            break;
                        case "EnableDebugLog":
                            if (bool.TryParse(value, out bool enableLog)) EnableDebugLog = enableLog;
                            break;
                        case "DefaultShippingTrackColumn": DefaultShippingTrackColumn = value; break;
                        case "DefaultShippingProductColumn": DefaultShippingProductColumn = value; break;
                        case "DefaultShippingNameColumn": DefaultShippingNameColumn = value; break;
                        case "DefaultBillTrackColumn": DefaultBillTrackColumn = value; break;
                        case "DefaultBillProductColumn": DefaultBillProductColumn = value; break;
                        case "DefaultBillNameColumn": DefaultBillNameColumn = value; break;
                        case "ProgressUpdateFrequency":
                            if (int.TryParse(value, out int freq)) ProgressUpdateFrequency = freq > 0 ? freq : 100;
                            break;
                        case "LogDirectory": LogDirectory = value; break;
                    }
                }
            }
            catch (Exception)
            {
                ResetToDefaults();
            }
        }

        public void Save()
        {
            try
            {
                var lines = new List<string>
                {
                    "# YY运单匹配工具设置文件",
                    "# 界面设置",
                    $"FontSize={FontSize}",
                    $"AutoScaleUI={AutoScaleUI}",
                    "",
                    "# 性能设置",
                    $"PerformanceMode={PerformanceMode}",
                    "",
                    "# WPS设置",
                    $"WPSPriority={WPSPriority}",
                    $"EnableDebugLog={EnableDebugLog}",
                    "",
                    "# 默认列设置",
                    $"DefaultShippingTrackColumn={DefaultShippingTrackColumn}",
                    $"DefaultShippingProductColumn={DefaultShippingProductColumn}",
                    $"DefaultShippingNameColumn={DefaultShippingNameColumn}",
                    $"DefaultBillTrackColumn={DefaultBillTrackColumn}",
                    $"DefaultBillProductColumn={DefaultBillProductColumn}",
                    $"DefaultBillNameColumn={DefaultBillNameColumn}",
                    "",
                    "# 高级设置",
                    $"ProgressUpdateFrequency={ProgressUpdateFrequency}",
                    $"LogDirectory={LogDirectory}"
                };

                File.WriteAllLines(ConfigPath, lines, System.Text.Encoding.UTF8);
            }
            catch (Exception)
            {
                // 保存失败不抛异常
            }
        }

        public void ResetToDefaults()
        {
            FontSize = 9;
            AutoScaleUI = true;
            PerformanceMode = PerformanceMode.UltraFast;
            WPSPriority = true;
            EnableDebugLog = true;
            DefaultShippingTrackColumn = "B";
            DefaultShippingProductColumn = "J";
            DefaultShippingNameColumn = "I";
            DefaultBillTrackColumn = "C";
            DefaultBillProductColumn = "Y";
            DefaultBillNameColumn = "Z";
            ProgressUpdateFrequency = 100;
            LogDirectory = "";
        }
    }

    /// <summary>
    /// 性能模式枚举
    /// </summary>
    public enum PerformanceMode
    {
        UltraFast = 0,    // 极速模式
        Balanced = 1,     // 平衡模式
        Compatible = 2    // 兼容模式
    }
}