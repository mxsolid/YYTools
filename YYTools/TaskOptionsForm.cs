using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 任务选项配置窗体
    /// </summary>
    public partial class TaskOptionsForm : Form
    {
        private AppSettings _settings;
        private bool _isInitializing = true;

        // 控件声明
        private GroupBox gbConcatenation;
        private Label lblDelimiter;
        private ComboBox cmbDelimiter;
        private CheckBox chkRemoveDuplicates;
        private ComboBox cmbSort;
        private Label lblSort;
        private GroupBox gbPerformance;
        private Label lblBatchSize;
        private NumericUpDown numBatchSize;
        private Label lblMaxPreviewRows;
        private NumericUpDown numMaxPreviewRows;
        private CheckBox chkEnableProgressReporting;
        private GroupBox gbSmartMatching;
        private CheckBox chkEnableSmartMatching;
        private CheckBox chkEnableExactMatchPriority;
        private TrackBar trkMinMatchScore;
        private Label lblMinMatchScore;
        private Label lblMinMatchScoreValue;
        private Button btnOK;
        private Button btnCancel;
        private Button btnReset;

        public TaskOptionsForm()
        {
            InitializeComponent();
            LoadSettings();
            ApplyUIEnhancement();
        }

        private void InitializeComponent()
        {
            this.Text = "任务选项配置";
            this.Size = new Size(500, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.BackColor = Color.White; // 确保背景是白色
            this.ForeColor = Color.Black; // 确保前景色是黑色

            // 分隔符配置组
            gbConcatenation = new GroupBox
            {
                Text = "拼接配置",
                Location = new Point(20, 20),
                Size = new Size(440, 120),
                BackColor = Color.White, // 确保分组框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            lblDelimiter = new Label
            {
                Text = "分隔符:",
                Location = new Point(20, 30),
                Size = new Size(80, 20),
                BackColor = Color.White, // 确保标签背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            cmbDelimiter = new ComboBox
            {
                Location = new Point(110, 28),
                Size = new Size(100, 20),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.White, // 确保下拉框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            chkRemoveDuplicates = new CheckBox
            {
                Text = "去除重复项",
                Location = new Point(20, 60),
                Size = new Size(120, 20),
                BackColor = Color.White, // 确保复选框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            lblSort = new Label
            {
                Text = "排序方式:",
                Location = new Point(220, 30),
                Size = new Size(80, 20),
                BackColor = Color.White, // 确保标签背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            cmbSort = new ComboBox
            {
                Location = new Point(310, 28),
                Size = new Size(100, 20),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.White, // 确保下拉框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            // 性能配置组
            gbPerformance = new GroupBox
            {
                Text = "性能配置",
                Location = new Point(20, 160),
                Size = new Size(440, 140),
                BackColor = Color.White, // 确保分组框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            lblBatchSize = new Label
            {
                Text = "批处理大小:",
                Location = new Point(20, 30),
                Size = new Size(100, 20),
                BackColor = Color.White, // 确保标签背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            numBatchSize = new NumericUpDown
            {
                Location = new Point(130, 28),
                Size = new Size(100, 20),
                Minimum = 100,
                Maximum = 10000,
                Increment = 100,
                BackColor = Color.White, // 确保数字输入框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            lblMaxPreviewRows = new Label
            {
                Text = "预览最大行数:",
                Location = new Point(20, 60),
                Size = new Size(100, 20),
                BackColor = Color.White, // 确保标签背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            numMaxPreviewRows = new NumericUpDown
            {
                Location = new Point(130, 58),
                Size = new Size(100, 20),
                Minimum = 10,
                Maximum = 1000,
                Increment = 10,
                BackColor = Color.White, // 确保数字输入框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            chkEnableProgressReporting = new CheckBox
            {
                Text = "启用进度报告",
                Location = new Point(20, 90),
                Size = new Size(120, 20),
                BackColor = Color.White, // 确保复选框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            // 智能匹配配置组
            gbSmartMatching = new GroupBox
            {
                Text = "智能匹配配置",
                Location = new Point(20, 320),
                Size = new Size(440, 140),
                BackColor = Color.White, // 确保分组框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            chkEnableSmartMatching = new CheckBox
            {
                Text = "启用智能匹配",
                Location = new Point(20, 30),
                Size = new Size(120, 20),
                BackColor = Color.White, // 确保复选框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            chkEnableExactMatchPriority = new CheckBox
            {
                Text = "完全匹配优先",
                Location = new Point(20, 60),
                Size = new Size(120, 20),
                BackColor = Color.White, // 确保复选框背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            lblMinMatchScore = new Label
            {
                Text = "最小匹配分数:",
                Location = new Point(20, 90),
                Size = new Size(100, 20),
                BackColor = Color.White, // 确保标签背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            trkMinMatchScore = new TrackBar
            {
                Location = new Point(130, 85),
                Size = new Size(200, 30),
                Minimum = 0,
                Maximum = 100,
                TickFrequency = 10,
                Value = 50,
                BackColor = Color.White, // 确保滑块背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            lblMinMatchScoreValue = new Label
            {
                Text = "0.5",
                Location = new Point(340, 90),
                Size = new Size(50, 20),
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.White, // 确保标签背景是白色
                ForeColor = Color.Black // 确保前景色是黑色
            };

            // 按钮
            btnOK = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Location = new Point(280, 500),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(0, 122, 204), // 蓝色按钮
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new Point(370, 500),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(100, 100, 100), // 灰色按钮
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnReset = new Button
            {
                Text = "重置默认",
                Location = new Point(20, 500),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(255, 152, 0), // 橙色按钮
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            // 添加控件到分组
            gbConcatenation.Controls.AddRange(new Control[] 
            {
                lblDelimiter, cmbDelimiter, chkRemoveDuplicates, lblSort, cmbSort
            });

            gbPerformance.Controls.AddRange(new Control[] 
            {
                lblBatchSize, numBatchSize, lblMaxPreviewRows, numMaxPreviewRows, chkEnableProgressReporting
            });

            gbSmartMatching.Controls.AddRange(new Control[] 
            {
                chkEnableSmartMatching, chkEnableExactMatchPriority, lblMinMatchScore, trkMinMatchScore, lblMinMatchScoreValue
            });

            // 添加所有控件到窗体
            this.Controls.AddRange(new Control[] 
            {
                gbConcatenation, gbPerformance, gbSmartMatching, btnOK, btnCancel, btnReset
            });

            // 绑定事件
            btnOK.Click += BtnOK_Click;
            btnReset.Click += BtnReset_Click;
            trkMinMatchScore.ValueChanged += TrkMinMatchScore_ValueChanged;
        }

        private void LoadSettings()
        {
            try
            {
                _settings = AppSettings.Instance;
                _isInitializing = true;

                // 加载分隔符选项
                cmbDelimiter.Items.Clear();
                cmbDelimiter.Items.AddRange(_settings.GetDelimiterOptions());
                cmbDelimiter.SelectedItem = _settings.ConcatenationDelimiter;

                // 加载排序选项
                cmbSort.Items.Clear();
                cmbSort.Items.AddRange(_settings.GetSortOptions());
                cmbSort.SelectedIndex = 0;

                // 加载性能配置
                numBatchSize.Value = _settings.BatchSize;
                numMaxPreviewRows.Value = _settings.MaxRowsForPreview;
                chkEnableProgressReporting.Checked = _settings.EnableProgressReporting;

                // 加载智能匹配配置
                chkEnableSmartMatching.Checked = _settings.EnableSmartMatching;
                chkEnableExactMatchPriority.Checked = _settings.EnableExactMatchPriority;
                trkMinMatchScore.Value = (int)(_settings.MinMatchScore * 100);

                // 加载其他配置
                chkRemoveDuplicates.Checked = _settings.RemoveDuplicateItems;

                _isInitializing = false;
            }
            catch (Exception ex)
            {
                Logger.LogError("加载任务选项设置失败", ex);
                MessageBox.Show($"加载设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveSettings()
        {
            try
            {
                if (_isInitializing) return;

                // 保存分隔符配置
                if (cmbDelimiter.SelectedItem != null)
                {
                    _settings.ConcatenationDelimiter = cmbDelimiter.SelectedItem.ToString();
                }

                // 保存排序配置
                if (cmbSort.SelectedIndex >= 0)
                {
                    // 这里可以根据需要保存排序配置
                }

                // 保存性能配置
                _settings.BatchSize = (int)numBatchSize.Value;
                _settings.MaxRowsForPreview = (int)numMaxPreviewRows.Value;
                _settings.EnableProgressReporting = chkEnableProgressReporting.Checked;

                // 保存智能匹配配置
                _settings.EnableSmartMatching = chkEnableSmartMatching.Checked;
                _settings.EnableExactMatchPriority = chkEnableExactMatchPriority.Checked;
                _settings.MinMatchScore = trkMinMatchScore.Value / 100.0;

                // 保存其他配置
                _settings.RemoveDuplicateItems = chkRemoveDuplicates.Checked;

                // 保存到文件
                _settings.Save();

                Logger.LogUserAction("保存任务选项配置", "配置已更新", "成功");
                MessageBox.Show("设置已保存", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.LogError("保存任务选项设置失败", ex);
                MessageBox.Show($"保存设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SaveSettings();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要重置所有设置为默认值吗？", "确认", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _settings.ResetToDefaults();
                    LoadSettings();
                    Logger.LogUserAction("重置任务选项配置", "所有设置已重置为默认值", "成功");
                    MessageBox.Show("设置已重置为默认值", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("重置任务选项设置失败", ex);
                MessageBox.Show($"重置设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TrkMinMatchScore_ValueChanged(object sender, EventArgs e)
        {
            if (!_isInitializing)
            {
                double value = trkMinMatchScore.Value / 100.0;
                lblMinMatchScoreValue.Text = value.ToString("F2");
            }
        }

        private void ApplyUIEnhancement()
        {
            try
            {
                // 应用UI美化
                UIEnhancer.EnhanceForm(this);
                UIEnhancer.EnhanceAllControls(this);

                // 美化按钮
                UIEnhancer.EnhanceButton(btnOK, UIEnhancer.ButtonStyle.Primary);
                UIEnhancer.EnhanceButton(btnCancel, UIEnhancer.ButtonStyle.Secondary);
                UIEnhancer.EnhanceButton(btnReset, UIEnhancer.ButtonStyle.Warning);
            }
            catch (Exception ex)
            {
                Logger.LogError("应用UI美化失败", ex);
            }
        }

        /// <summary>
        /// 显示任务选项配置窗体
        /// </summary>
        public static void ShowTaskOptions(IWin32Window owner = null)
        {
            try
            {
                var form = new TaskOptionsForm();
                form.ShowDialog(owner);
            }
            catch (Exception ex)
            {
                Logger.LogError("显示任务选项配置窗体失败", ex);
                MessageBox.Show("显示任务选项配置失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}