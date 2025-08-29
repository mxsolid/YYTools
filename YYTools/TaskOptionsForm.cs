using System;
using System.Linq;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 任务选项配置窗体 - 采用TabControl重构界面
    /// </summary>
    public partial class TaskOptionsForm : Form
    {
        private AppSettings _settings;
        private bool _isInitializing = true;

        public TaskOptionsForm()
        {
            InitializeComponent(); // 由 Designer.cs 文件提供
            // 初始化时加载当前设置
            LoadSettings();
        }

        /// <summary>
        /// 从 AppSettings 单例加载设置并填充到UI控件中
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                _settings = AppSettings.Instance;
                _isInitializing = true; // 标记正在初始化，防止触发不必要的事件

                // === 拼接设置 Tab ===
                // 加载分隔符选项
                cmbDelimiter.Items.Clear();
                cmbDelimiter.Items.AddRange(_settings.GetDelimiterOptions());
                cmbDelimiter.SelectedItem = _settings.ConcatenationDelimiter;

                // 加载排序选项 (这里简化处理，您可以根据需要扩展)
                cmbSort.Items.Clear();
                cmbSort.Items.AddRange(new string[] { "默认", "升序", "降序" });
                cmbSort.SelectedIndex = 0; // 默认为“默认”

                // 去除重复项
                chkRemoveDuplicates.Checked = _settings.RemoveDuplicateItems;

                // === 性能与预览 Tab ===
                numBatchSize.Value = Math.Max(numBatchSize.Minimum, Math.Min(numBatchSize.Maximum, _settings.BatchSize));
                numMaxPreviewRows.Value = Math.Max(numMaxPreviewRows.Minimum, Math.Min(numMaxPreviewRows.Maximum, _settings.MaxRowsForPreview));
                chkEnableProgressReporting.Checked = _settings.EnableProgressReporting;
                chkEnableColumnDataPreview.Checked = _settings.EnableColumnDataPreview;
                chkEnableWritePreview.Checked = _settings.EnableWritePreview;

                // 加载预览行数选项
                cmbPreviewRows.Items.Clear();
                cmbPreviewRows.Items.AddRange(_settings.GetPreviewRowOptions().Cast<object>().ToArray());
                cmbPreviewRows.SelectedItem = _settings.PreviewParseRows;
                
                // === 智能匹配 Tab ===
                chkEnableSmartMatching.Checked = _settings.EnableSmartMatching;
                chkEnableExactMatchPriority.Checked = _settings.EnableExactMatchPriority;
                // 设置滑块值，并确保在0-100范围内
                int scoreValue = (int)(_settings.MinMatchScore * 100);
                trkMinMatchScore.Value = Math.Max(trkMinMatchScore.Minimum, Math.Min(trkMinMatchScore.Maximum, scoreValue));
                // 手动更新一次标签文本
                lblMinMatchScoreValue.Text = (_settings.MinMatchScore).ToString("F2");


            }
            catch (Exception ex)
            {
                Logger.LogError("加载任务选项设置失败", ex);
                MessageBox.Show($"加载设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isInitializing = false; // 初始化完成
            }
        }

        /// <summary>
        /// 将UI控件中的设置保存回 AppSettings 单例
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                if (_isInitializing) return;

                // === 保存拼接设置 ===
                if (cmbDelimiter.SelectedItem != null)
                {
                    _settings.ConcatenationDelimiter = cmbDelimiter.SelectedItem.ToString();
                }
                _settings.RemoveDuplicateItems = chkRemoveDuplicates.Checked;
                // 根据需要保存排序设置...

                // === 保存性能与预览设置 ===
                _settings.BatchSize = (int)numBatchSize.Value;
                _settings.MaxRowsForPreview = (int)numMaxPreviewRows.Value;
                _settings.EnableProgressReporting = chkEnableProgressReporting.Checked;
                _settings.EnableColumnDataPreview = chkEnableColumnDataPreview.Checked;
                _settings.EnableWritePreview = chkEnableWritePreview.Checked;
                if (cmbPreviewRows.SelectedItem != null)
                {
                    _settings.PreviewParseRows = (int)cmbPreviewRows.SelectedItem;
                }

                // === 保存智能匹配设置 ===
                _settings.EnableSmartMatching = chkEnableSmartMatching.Checked;
                _settings.EnableExactMatchPriority = chkEnableExactMatchPriority.Checked;
                _settings.MinMatchScore = trkMinMatchScore.Value / 100.0;

                // 持久化保存到文件
                _settings.Save();
                Logger.LogUserAction("保存任务选项配置成功");
            }
            catch (Exception ex)
            {
                Logger.LogError("保存任务选项设置失败", ex);
                MessageBox.Show($"保存设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- 事件处理程序 ---

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

        private void btnReset_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("您确定要将所有任务选项重置为默认值吗？", "确认重置",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _settings.ResetToDefaults(); // 调用重置方法
                    LoadSettings(); // 重新加载默认设置到界面
                    Logger.LogUserAction("任务选项已重置为默认值");
                    MessageBox.Show("设置已成功重置为默认值。", "操作成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("重置任务选项设置失败", ex);
                MessageBox.Show($"重置设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void trkMinMatchScore_ValueChanged(object sender, EventArgs e)
        {
            // 只有在非初始化阶段才更新，避免加载时触发
            if (!_isInitializing)
            {
                double value = trkMinMatchScore.Value / 100.0;
                lblMinMatchScoreValue.Text = value.ToString("F2");
            }
        }
        
        /// <summary>
        /// [静态方法] 用于从外部调用显示本窗体
        /// </summary>
        public static void ShowTaskOptions(IWin32Window owner = null)
        {
            try
            {
                using (var form = new TaskOptionsForm())
                {
                    form.ShowDialog(owner);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("显示任务选项配置窗体失败", ex);
                MessageBox.Show("无法打开任务选项配置窗体：" + ex.Message, "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}