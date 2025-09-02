using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    public partial class SettingsForm : Form
    {
        private AppSettings settings;

        public SettingsForm()
        {
            InitializeComponent();
            this.AutoScaleMode = AutoScaleMode.Font;
            settings = AppSettings.Instance;
            LoadSettings();
            ApplyCurrentFontSettings();
        }

        private void ApplyCurrentFontSettings()
        {
            try
            {
                Font currentFont = new Font("微软雅黑", settings.FontSize, FontStyle.Regular);
                this.Font = currentFont;
                this.PerformAutoScale();
            }
            catch { }
        }

        private void LoadSettings()
        {
            try
            {
                numFontSize.Value = settings.FontSize;
                chkAutoScale.Checked = settings.AutoScaleUI;
                txtLogDirectory.Text = settings.LogDirectory;
                
                cmbMaxThreads.Items.Clear();
                for (int i = 1; i <= Environment.ProcessorCount; i++)
                {
                    cmbMaxThreads.Items.Add(i);
                }
                cmbMaxThreads.SelectedItem = settings.MaxThreads > 0 ? settings.MaxThreads : Environment.ProcessorCount;
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
                settings.LogDirectory = txtLogDirectory.Text;
                settings.MaxThreads = (int)cmbMaxThreads.SelectedItem;
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

        private void btnResetDefaults_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定要重置为默认设置吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                settings.ResetToDefaults();
                LoadSettings();
                MessageBox.Show("已重置为默认设置！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnBrowseLog_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "请选择日志文件存储目录";
                fbd.SelectedPath = txtLogDirectory.Text;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtLogDirectory.Text = fbd.SelectedPath;
                }
            }
        }
    }
}