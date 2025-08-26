using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 设置窗体 - 已修复DPI和字体缩放问题
    /// </summary>
    public partial class SettingsForm : Form
    {
        private AppSettings settings;
        
        public SettingsForm()
        {
            InitializeComponent();
            
            // 关键：将AutoScaleMode设置为Font，以便根据字体大小自动调整布局
            this.AutoScaleMode = AutoScaleMode.Font;
            
            settings = AppSettings.Instance;
            LoadSettings();
            
            // 在加载设置后，应用当前字体并强制重新缩放
            ApplyCurrentFontSettings();
        }
        
        /// <summary>
        /// 应用当前字体设置到设置窗体，并强制重新缩放
        /// </summary>
        private void ApplyCurrentFontSettings()
        {
            try
            {
                Font currentFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                this.Font = currentFont; // 应用基础字体
                ApplyFontToAllControls(this, currentFont);

                // 关键：强制窗体根据新字体重新计算布局
                this.PerformAutoScale();
            }
            catch
            {
                // 字体应用失败时使用默认字体
            }
        }
        
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
                MessageBox.Show("加载设置失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void SaveSettings()
        {
            try
            {
                settings.FontSize = (int)numFontSize.Value;
                settings.AutoScaleUI = chkAutoScale.Checked;
                settings.PerformanceMode = (PerformanceMode)cmbPerformanceMode.SelectedIndex;
                settings.WPSPriority = chkWPSPriority.Checked;
                settings.EnableDebugLog = chkEnableDebugLog.Checked;
                settings.DefaultShippingTrackColumn = txtShippingTrack.Text.Trim().ToUpper();
                settings.DefaultShippingProductColumn = txtShippingProduct.Text.Trim().ToUpper();
                settings.DefaultShippingNameColumn = txtShippingName.Text.Trim().ToUpper();
                settings.DefaultBillTrackColumn = txtBillTrack.Text.Trim().ToUpper();
                settings.DefaultBillProductColumn = txtBillProduct.Text.Trim().ToUpper();
                settings.DefaultBillNameColumn = txtBillName.Text.Trim().ToUpper();
                settings.ProgressUpdateFrequency = (int)numProgressFreq.Value;
                settings.LogDirectory = txtLogDirectory.Text.Trim();
                
                settings.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存设置失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            MessageBox.Show("设置已应用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            if (MessageBox.Show("确定要重置为默认设置吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                settings.ResetToDefaults();
                LoadSettings();
                MessageBox.Show("已重置为默认设置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}