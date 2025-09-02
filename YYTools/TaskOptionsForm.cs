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

        // 标志位，用于通知主窗体设置变更后是否需要刷新数据
        public bool SettingsChangedNeedRefresh { get; private set; } = false;

        public TaskOptionsForm()
        {
            InitializeComponent(); // 由 Designer.cs 文件提供
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
                _isInitializing = true;

                // === 拼接设置 Tab ===
                cmbDelimiter.Items.Clear();
                // **修正点**: 直接从 Constants 类加载选项，而不是通过 AppSettings 实例的方法
                cmbDelimiter.Items.AddRange(Constants.DelimiterOptions);
                cmbDelimiter.SelectedItem = _settings.ConcatenationDelimiter;

                cmbSort.Items.Clear();
                cmbSort.Items.AddRange(new string[] { "默认", "升序", "降序" });
                switch (_settings.SortOption)
                {
                    case SortOption.Asc:
                        cmbSort.SelectedItem = "升序";
                        break;
                    case SortOption.Desc:
                        cmbSort.SelectedItem = "降序";
                        break;
                    default:
                        cmbSort.SelectedItem = "默认";
                        break;
                }

                chkRemoveDuplicates.Checked = _settings.RemoveDuplicateItems;

                // === 性能与预览 Tab ===
                chkEnableColumnPreview.Checked = _settings.EnableColumnDataPreview;
                chkEnableWritePreview.Checked = _settings.EnableWritePreview;
                numBatchSize.Value = Math.Max(numBatchSize.Minimum, Math.Min(numBatchSize.Maximum, _settings.BatchSize));
                numMaxPreviewRows.Value = Math.Max(numMaxPreviewRows.Minimum, Math.Min(numMaxPreviewRows.Maximum, _settings.MaxRowsForPreview));
                chkEnableProgressReporting.Checked = _settings.EnableProgressReporting;

                cmbPreviewRows.Items.Clear();
                // **修正点**: 直接从 Constants 类加载选项
                cmbPreviewRows.Items.AddRange(Constants.PreviewRowOptions.Cast<object>().ToArray());
                cmbPreviewRows.SelectedItem = _settings.PreviewParseRows;
                
                // === 智能匹配 Tab ===
                chkEnableSmartMatching.Checked = _settings.EnableSmartMatching;
                chkEnableExactMatchPriority.Checked = _settings.EnableExactMatchPriority;
                int scoreValue = (int)(_settings.MinMatchScore * 100);
                trkMinMatchScore.Value = Math.Max(trkMinMatchScore.Minimum, Math.Min(trkMinMatchScore.Maximum, scoreValue));
                lblMinMatchScoreValue.Text = (_settings.MinMatchScore).ToString("F2");
            }
            catch (Exception ex)
            {
                Logger.LogError("加载任务选项设置失败", ex);
                MessageBox.Show($"加载设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isInitializing = false;
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

                // 检查关键设置（列预览）是否有变动，如有则通知主窗体刷新
                if (_settings.EnableColumnDataPreview != chkEnableColumnPreview.Checked)
                {
                    SettingsChangedNeedRefresh = true;
                }

                // === 保存拼接设置 ===
                if (cmbDelimiter.SelectedItem != null)
                {
                    _settings.ConcatenationDelimiter = cmbDelimiter.SelectedItem.ToString();
                }
                _settings.RemoveDuplicateItems = chkRemoveDuplicates.Checked;
                
                switch (cmbSort.SelectedItem.ToString())
                {
                    case "升序":
                        _settings.SortOption = SortOption.Asc;
                        break;
                    case "降序":
                        _settings.SortOption = SortOption.Desc;
                        break;
                    default:
                        _settings.SortOption = SortOption.None;
                        break;
                }
                
                // === 保存性能与预览设置 ===
                _settings.EnableColumnDataPreview = chkEnableColumnPreview.Checked;
                _settings.EnableWritePreview = chkEnableWritePreview.Checked;
                _settings.BatchSize = (int)numBatchSize.Value;
                _settings.MaxRowsForPreview = (int)numMaxPreviewRows.Value;
                _settings.EnableProgressReporting = chkEnableProgressReporting.Checked;
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
                    _settings.ResetToDefaults();
                    LoadSettings();
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
                    // 将 ShowDialog 的结果传递给主窗体，以便判断是否需要刷新
                    if (form.ShowDialog(owner) == DialogResult.OK && form.SettingsChangedNeedRefresh)
                    {
                        if (owner is MatchForm matchForm)
                        {
                            matchForm.TriggerRefresh();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("显示任务选项配置窗体失败", ex);
                // 抛出异常，让调用方（MatchForm）来处理弹窗
                throw;
            }
        }
    }
}