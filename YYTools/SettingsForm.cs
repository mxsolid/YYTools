using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 设置窗体 - WPS优先配置
    /// </summary>
    public partial class SettingsForm : Form
    {
        private AppSettings settings;
        
        public SettingsForm()
        {
            InitializeComponent();
            
            // 高分辨率显示优化
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(6F, 12F);
            
            // 适配高清屏幕
            if (Environment.OSVersion.Version.Major >= 6) // Vista及以上
            {
                SetProcessDPIAware(); // 启用DPI感知
            }
            
            settings = AppSettings.Instance;
            LoadSettings();
            
            // 应用当前字体设置到设置窗体
            ApplyCurrentFontSettings();
        }
        
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
        
        /// <summary>
        /// 应用当前字体设置到设置窗体
        /// </summary>
        private void ApplyCurrentFontSettings()
        {
            try
            {
                Font currentFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                ApplyFontToAllControls(this, currentFont);
            }
            catch
            {
                // 字体应用失败时使用默认字体
            }
        }
        
        /// <summary>
        /// 递归应用字体到所有控件
        /// </summary>
        private void ApplyFontToAllControls(Control parent, Font font)
        {
            foreach (Control control in parent.Controls)
            {
                control.Font = font;
                if (control.HasChildren)
                {
                    ApplyFontToAllControls(control, font);
                }
            }
        }
        
        private void LoadSettings()
        {
            try
            {
                // 字体设置
                numFontSize.Value = settings.FontSize;
                chkAutoScale.Checked = settings.AutoScaleUI;
                
                // 性能模式
                cmbPerformanceMode.SelectedIndex = (int)settings.PerformanceMode;
                
                // WPS优先设置
                chkWPSPriority.Checked = settings.WPSPriority;
                chkEnableDebugLog.Checked = settings.EnableDebugLog;
                
                // 默认列设置
                txtShippingTrack.Text = settings.DefaultShippingTrackColumn;
                txtShippingProduct.Text = settings.DefaultShippingProductColumn;
                txtShippingName.Text = settings.DefaultShippingNameColumn;
                txtBillTrack.Text = settings.DefaultBillTrackColumn;
                txtBillProduct.Text = settings.DefaultBillProductColumn;
                txtBillName.Text = settings.DefaultBillNameColumn;
                
                // 高级设置
                numProgressFreq.Value = settings.ProgressUpdateFrequency;
                txtLogDirectory.Text = settings.LogDirectory;
            }
            catch (Exception ex)
            {
                MessageBox.Show("加载设置失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void SaveSettings()
        {
            try
            {
                // 字体设置
                settings.FontSize = (int)numFontSize.Value;
                settings.AutoScaleUI = chkAutoScale.Checked;
                
                // 性能模式
                settings.PerformanceMode = (PerformanceMode)cmbPerformanceMode.SelectedIndex;
                
                // WPS优先设置
                settings.WPSPriority = chkWPSPriority.Checked;
                settings.EnableDebugLog = chkEnableDebugLog.Checked;
                
                // 默认列设置
                settings.DefaultShippingTrackColumn = txtShippingTrack.Text.Trim();
                settings.DefaultShippingProductColumn = txtShippingProduct.Text.Trim();
                settings.DefaultShippingNameColumn = txtShippingName.Text.Trim();
                settings.DefaultBillTrackColumn = txtBillTrack.Text.Trim();
                settings.DefaultBillProductColumn = txtBillProduct.Text.Trim();
                settings.DefaultBillNameColumn = txtBillName.Text.Trim();
                
                // 高级设置
                settings.ProgressUpdateFrequency = (int)numProgressFreq.Value;
                settings.LogDirectory = txtLogDirectory.Text.Trim();
                
                // 保存到配置文件
                settings.Save();
                
                MessageBox.Show("设置已保存！", "成功", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存设置失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void btnOK_Click(object sender, EventArgs e)
        {
            SaveSettings();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
        
        private void btnApply_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }
        
        private void btnBrowseLog_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择日志目录";
                dialog.SelectedPath = txtLogDirectory.Text;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtLogDirectory.Text = dialog.SelectedPath;
                }
            }
        }
        
        private void btnResetDefaults_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定要重置为默认设置吗？", "确认", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                settings.ResetToDefaults();
                LoadSettings();
                MessageBox.Show("已重置为默认设置！", "提示", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
    
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
                string folder = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "YYTools");
                
                if (!System.IO.Directory.Exists(folder))
                    System.IO.Directory.CreateDirectory(folder);
                
                return System.IO.Path.Combine(folder, "settings.ini");
            }
        }
        
        public void Load()
        {
            try
            {
                if (!System.IO.File.Exists(ConfigPath))
                {
                    ResetToDefaults();
                    Save();
                    return;
                }
                
                string[] lines = System.IO.File.ReadAllLines(ConfigPath, System.Text.Encoding.UTF8);
                foreach (string line in lines)
                {
                    if (string.IsNullOrEmpty(line) || line.StartsWith("#")) continue;
                    
                    string[] parts = line.Split('=');
                    if (parts.Length != 2) continue;
                    
                    string key = parts[0].Trim();
                    string value = parts[1].Trim();
                    
                    switch (key)
                    {
                        case "FontSize":
                            int fontSize;
                            int.TryParse(value, out fontSize);
                            FontSize = fontSize > 0 ? fontSize : 9;
                            break;
                        case "AutoScaleUI":
                            bool autoScale;
                            bool.TryParse(value, out autoScale);
                            AutoScaleUI = autoScale;
                            break;
                        case "PerformanceMode":
                            try
                            {
                                PerformanceMode = (PerformanceMode)Enum.Parse(typeof(PerformanceMode), value);
                            }
                            catch
                            {
                                PerformanceMode = PerformanceMode.UltraFast;
                            }
                            break;
                        case "WPSPriority":
                            bool wpsPriority;
                            bool.TryParse(value, out wpsPriority);
                            WPSPriority = wpsPriority;
                            break;
                        case "EnableDebugLog":
                            bool enableLog;
                            bool.TryParse(value, out enableLog);
                            EnableDebugLog = enableLog;
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
                        case "ProgressUpdateFrequency":
                            int freq;
                            int.TryParse(value, out freq);
                            ProgressUpdateFrequency = freq > 0 ? freq : 100;
                            break;
                        case "LogDirectory":
                            LogDirectory = value;
                            break;
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
                var lines = new string[]
                {
                    "# YY运单匹配工具设置文件",
                    "# 界面设置",
                    "FontSize=" + FontSize,
                    "AutoScaleUI=" + AutoScaleUI,
                    "",
                    "# 性能设置",
                    "PerformanceMode=" + PerformanceMode,
                    "",
                    "# WPS设置",
                    "WPSPriority=" + WPSPriority,
                    "EnableDebugLog=" + EnableDebugLog,
                    "",
                    "# 默认列设置",
                    "DefaultShippingTrackColumn=" + DefaultShippingTrackColumn,
                    "DefaultShippingProductColumn=" + DefaultShippingProductColumn,
                    "DefaultShippingNameColumn=" + DefaultShippingNameColumn,
                    "DefaultBillTrackColumn=" + DefaultBillTrackColumn,
                    "DefaultBillProductColumn=" + DefaultBillProductColumn,
                    "DefaultBillNameColumn=" + DefaultBillNameColumn,
                    "",
                    "# 高级设置",
                    "ProgressUpdateFrequency=" + ProgressUpdateFrequency,
                    "LogDirectory=" + LogDirectory
                };
                
                System.IO.File.WriteAllLines(ConfigPath, lines, System.Text.Encoding.UTF8);
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
