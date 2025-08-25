using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 设置窗体 - 支持性能模式、字体、默认值配置
    /// </summary>
    public partial class SettingsForm : Form
    {
        private AppSettings settings;
        
        public SettingsForm()
        {
            InitializeComponent();
            settings = AppSettings.Instance;
            LoadSettings();
        }

        /// <summary>
        /// 加载当前设置
        /// </summary>
        private void LoadSettings()
        {
            // 性能模式
            cmbPerformanceMode.SelectedIndex = (int)settings.PerformanceMode;
            
            // 字体设置
            nudFontSize.Value = (decimal)settings.FontSize;
            chkAutoScale.Checked = settings.AutoScaleUI;
            
            // 默认值设置
            txtDefaultShippingTrack.Text = settings.DefaultShippingTrackColumn;
            txtDefaultShippingProduct.Text = settings.DefaultShippingProductColumn;
            txtDefaultShippingName.Text = settings.DefaultShippingNameColumn;
            txtDefaultBillTrack.Text = settings.DefaultBillTrackColumn;
            txtDefaultBillProduct.Text = settings.DefaultBillProductColumn;
            txtDefaultBillName.Text = settings.DefaultBillNameColumn;
            
            // 其他设置
            txtLogDirectory.Text = settings.LogDirectory;
            chkAutoSelectSheets.Checked = settings.AutoSelectSheets;
            nudProgressUpdateInterval.Value = settings.ProgressUpdateInterval;
        }

        /// <summary>
        /// 保存设置按钮
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // 性能模式
                settings.PerformanceMode = (PerformanceMode)cmbPerformanceMode.SelectedIndex;
                
                // 字体设置
                settings.FontSize = (float)nudFontSize.Value;
                settings.AutoScaleUI = chkAutoScale.Checked;
                
                // 默认值设置
                settings.DefaultShippingTrackColumn = txtDefaultShippingTrack.Text.Trim();
                settings.DefaultShippingProductColumn = txtDefaultShippingProduct.Text.Trim();
                settings.DefaultShippingNameColumn = txtDefaultShippingName.Text.Trim();
                settings.DefaultBillTrackColumn = txtDefaultBillTrack.Text.Trim();
                settings.DefaultBillProductColumn = txtDefaultBillProduct.Text.Trim();
                settings.DefaultBillNameColumn = txtDefaultBillName.Text.Trim();
                
                // 其他设置
                settings.LogDirectory = txtLogDirectory.Text.Trim();
                settings.AutoSelectSheets = chkAutoSelectSheets.Checked;
                settings.ProgressUpdateInterval = (int)nudProgressUpdateInterval.Value;
                
                // 保存到文件
                settings.Save();
                
                MessageBox.Show("设置保存成功！", "提示", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("保存设置失败：{0}", ex.Message), "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 重置默认设置
        /// </summary>
        private void btnReset_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("确定要重置为默认设置吗？", "确认", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                settings.ResetToDefaults();
                LoadSettings();
            }
        }

        /// <summary>
        /// 浏览日志目录
        /// </summary>
        private void btnBrowseLog_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择日志存储目录";
                dialog.SelectedPath = txtLogDirectory.Text;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtLogDirectory.Text = dialog.SelectedPath;
                }
            }
        }

        /// <summary>
        /// 取消按钮
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        /// <summary>
        /// 性能模式改变事件
        /// </summary>
        private void cmbPerformanceMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            PerformanceMode mode = (PerformanceMode)cmbPerformanceMode.SelectedIndex;
            
            switch (mode)
            {
                case PerformanceMode.UltraFast:
                    lblPerformanceDesc.Text = "极速模式：最高性能，适用于高配置机器（推荐）";
                    break;
                case PerformanceMode.Balanced:
                    lblPerformanceDesc.Text = "平衡模式：兼顾性能和兼容性，适用于大多数机器";
                    break;
                case PerformanceMode.Compatible:
                    lblPerformanceDesc.Text = "兼容模式：最佳兼容性，适用于低配置或老旧机器";
                    break;
            }
        }
    }

    /// <summary>
    /// 性能模式枚举
    /// </summary>
    public enum PerformanceMode
    {
        UltraFast = 0,      // 极速模式
        Balanced = 1,       // 平衡模式  
        Compatible = 2      // 兼容模式
    }

    /// <summary>
    /// 应用设置类 - 单例模式
    /// </summary>
    public class AppSettings
    {
        private static AppSettings instance;
        private static readonly object lockObject = new object();
        
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

        private string settingsPath;

        private AppSettings()
        {
            string appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "YYTools");
            
            if (!Directory.Exists(appDataPath))
                Directory.CreateDirectory(appDataPath);
                
            settingsPath = Path.Combine(appDataPath, "settings.ini");
            ResetToDefaults();
        }

        // 性能设置
        public PerformanceMode PerformanceMode { get; set; }
        
        // UI设置
        public float FontSize { get; set; }
        public bool AutoScaleUI { get; set; }
        
        // 默认列设置
        public string DefaultShippingTrackColumn { get; set; }
        public string DefaultShippingProductColumn { get; set; }
        public string DefaultShippingNameColumn { get; set; }
        public string DefaultBillTrackColumn { get; set; }
        public string DefaultBillProductColumn { get; set; }
        public string DefaultBillNameColumn { get; set; }
        
        // 其他设置
        public string LogDirectory { get; set; }
        public bool AutoSelectSheets { get; set; }
        public int ProgressUpdateInterval { get; set; }

        /// <summary>
        /// 重置为默认值
        /// </summary>
        public void ResetToDefaults()
        {
            PerformanceMode = PerformanceMode.UltraFast;
            FontSize = 9F;
            AutoScaleUI = true;
            
            DefaultShippingTrackColumn = "B";
            DefaultShippingProductColumn = "J";
            DefaultShippingNameColumn = "I";
            DefaultBillTrackColumn = "C";
            DefaultBillProductColumn = "Y";
            DefaultBillNameColumn = "Z";
            
            LogDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "YYTools", "Logs");
            AutoSelectSheets = true;
            ProgressUpdateInterval = 500;
        }

        /// <summary>
        /// 从文件加载设置
        /// </summary>
        public void Load()
        {
            try
            {
                if (File.Exists(settingsPath))
                {
                    string[] lines = File.ReadAllLines(settingsPath);
                    foreach (string line in lines)
                    {
                        if (line.Contains("="))
                        {
                            string[] parts = line.Split('=');
                            if (parts.Length == 2)
                            {
                                string key = parts[0].Trim();
                                string value = parts[1].Trim();
                                
                                switch (key)
                                {
                                    case "PerformanceMode":
                                        PerformanceMode = (PerformanceMode)int.Parse(value);
                                        break;
                                    case "FontSize":
                                        FontSize = float.Parse(value);
                                        break;
                                    case "AutoScaleUI":
                                        AutoScaleUI = bool.Parse(value);
                                        break;
                                    case "DefaultShippingTrackColumn":
                                        DefaultShippingTrackColumn = value;
                                        break;
                                    case "DefaultShippingProductColumn":
                                        DefaultShippingProductColumn = value;
                                        break;
                                    case "DefaultShippingNameColumn":
                                        DefaultShippingNameColumn = value;
                                        break;
                                    case "DefaultBillTrackColumn":
                                        DefaultBillTrackColumn = value;
                                        break;
                                    case "DefaultBillProductColumn":
                                        DefaultBillProductColumn = value;
                                        break;
                                    case "DefaultBillNameColumn":
                                        DefaultBillNameColumn = value;
                                        break;
                                    case "LogDirectory":
                                        LogDirectory = value;
                                        break;
                                    case "AutoSelectSheets":
                                        AutoSelectSheets = bool.Parse(value);
                                        break;
                                    case "ProgressUpdateInterval":
                                        ProgressUpdateInterval = int.Parse(value);
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // 加载失败则使用默认值
                ResetToDefaults();
            }
        }

        /// <summary>
        /// 保存设置到文件
        /// </summary>
        public void Save()
        {
            try
            {
                string[] lines = new string[]
                {
                    string.Format("PerformanceMode={0}", (int)PerformanceMode),
                    string.Format("FontSize={0}", FontSize),
                    string.Format("AutoScaleUI={0}", AutoScaleUI),
                    string.Format("DefaultShippingTrackColumn={0}", DefaultShippingTrackColumn),
                    string.Format("DefaultShippingProductColumn={0}", DefaultShippingProductColumn),
                    string.Format("DefaultShippingNameColumn={0}", DefaultShippingNameColumn),
                    string.Format("DefaultBillTrackColumn={0}", DefaultBillTrackColumn),
                    string.Format("DefaultBillProductColumn={0}", DefaultBillProductColumn),
                    string.Format("DefaultBillNameColumn={0}", DefaultBillNameColumn),
                    string.Format("LogDirectory={0}", LogDirectory),
                    string.Format("AutoSelectSheets={0}", AutoSelectSheets),
                    string.Format("ProgressUpdateInterval={0}", ProgressUpdateInterval)
                };
                
                File.WriteAllLines(settingsPath, lines);
            }
            catch (Exception ex)
            {
                throw new Exception("保存设置文件失败: " + ex.Message);
            }
        }
    }
} 