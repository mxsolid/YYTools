using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;

namespace YYTools
{
    /// <summary>
    /// UI控件美化器
    /// </summary>
    public static class UIEnhancer
    {
        #region 颜色主题

        /// <summary>
        /// 默认颜色主题
        /// </summary>
        public static class DefaultTheme
        {
            public static Color PrimaryColor = Color.FromArgb(0, 122, 204);
            public static Color SecondaryColor = Color.FromArgb(45, 45, 48);
            public static Color AccentColor = Color.FromArgb(0, 153, 204);
            public static Color SuccessColor = Color.FromArgb(76, 175, 80);
            public static Color WarningColor = Color.FromArgb(255, 152, 0);
            public static Color ErrorColor = Color.FromArgb(244, 67, 54);
            public static Color BackgroundColor = Color.FromArgb(30, 30, 30);
            public static Color SurfaceColor = Color.FromArgb(45, 45, 48);
            public static Color TextColor = Color.FromArgb(241, 241, 241);
            public static Color TextSecondaryColor = Color.FromArgb(200, 200, 200);
            public static Color BorderColor = Color.FromArgb(60, 60, 60);
            public static Color HighlightColor = Color.FromArgb(0, 122, 204);
        }

        /// <summary>
        /// 浅色主题
        /// </summary>
        public static class LightTheme
        {
            public static Color PrimaryColor = Color.FromArgb(0, 122, 204);
            public static Color SecondaryColor = Color.FromArgb(240, 240, 240);
            public static Color AccentColor = Color.FromArgb(0, 153, 204);
            public static Color SuccessColor = Color.FromArgb(76, 175, 80);
            public static Color WarningColor = Color.FromArgb(255, 152, 0);
            public static Color ErrorColor = Color.FromArgb(244, 67, 54);
            public static Color BackgroundColor = Color.FromArgb(255, 255, 255);
            public static Color SurfaceColor = Color.FromArgb(248, 248, 248);
            public static Color TextColor = Color.FromArgb(30, 30, 30);
            public static Color TextSecondaryColor = Color.FromArgb(100, 100, 100);
            public static Color BorderColor = Color.FromArgb(200, 200, 200);
            public static Color HighlightColor = Color.FromArgb(0, 122, 204);
        }

        #endregion

        #region 美化方法

        /// <summary>
        /// 美化ComboBox控件
        /// </summary>
        public static void EnhanceComboBox(ComboBox comboBox, bool useModernStyle = true)
        {
            try
            {
                if (useModernStyle)
                {
                    comboBox.FlatStyle = FlatStyle.Flat;
                    comboBox.BackColor = DefaultTheme.SurfaceColor;
                    comboBox.ForeColor = DefaultTheme.TextColor;
                    comboBox.BorderStyle = BorderStyle.FixedSingle;
                    
                    // 设置字体
                    comboBox.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                    
                    // 设置下拉箭头颜色
                    comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                else
                {
                    comboBox.FlatStyle = FlatStyle.Standard;
                    comboBox.BackColor = SystemColors.Window;
                    comboBox.ForeColor = SystemColors.WindowText;
                    comboBox.BorderStyle = BorderStyle.Fixed3D;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化ComboBox失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化Button控件
        /// </summary>
        public static void EnhanceButton(Button button, ButtonStyle style = ButtonStyle.Primary)
        {
            try
            {
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderSize = 0;
                button.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                button.Cursor = Cursors.Hand;

                switch (style)
                {
                    case ButtonStyle.Primary:
                        button.BackColor = DefaultTheme.PrimaryColor;
                        button.ForeColor = Color.White;
                        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                            Math.Min(255, DefaultTheme.PrimaryColor.R + 20),
                            Math.Min(255, DefaultTheme.PrimaryColor.G + 20),
                            Math.Min(255, DefaultTheme.PrimaryColor.B + 20)
                        );
                        break;
                    case ButtonStyle.Secondary:
                        button.BackColor = DefaultTheme.SecondaryColor;
                        button.ForeColor = DefaultTheme.TextColor;
                        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                            Math.Min(255, DefaultTheme.SecondaryColor.R + 20),
                            Math.Min(255, DefaultTheme.SecondaryColor.G + 20),
                            Math.Min(255, DefaultTheme.SecondaryColor.B + 20)
                        );
                        break;
                    case ButtonStyle.Success:
                        button.BackColor = DefaultTheme.SuccessColor;
                        button.ForeColor = Color.White;
                        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                            Math.Min(255, DefaultTheme.SuccessColor.R + 20),
                            Math.Min(255, DefaultTheme.SuccessColor.G + 20),
                            Math.Min(255, DefaultTheme.SuccessColor.B + 20)
                        );
                        break;
                    case ButtonStyle.Warning:
                        button.BackColor = DefaultTheme.WarningColor;
                        button.ForeColor = Color.White;
                        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                            Math.Min(255, DefaultTheme.WarningColor.R + 20),
                            Math.Min(255, DefaultTheme.WarningColor.G + 20),
                            Math.Min(255, DefaultTheme.WarningColor.B + 20)
                        );
                        break;
                    case ButtonStyle.Danger:
                        button.BackColor = DefaultTheme.ErrorColor;
                        button.ForeColor = Color.White;
                        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                            Math.Min(255, DefaultTheme.ErrorColor.R + 20),
                            Math.Min(255, DefaultTheme.ErrorColor.G + 20),
                            Math.Min(255, DefaultTheme.ErrorColor.B + 20)
                        );
                        break;
                }

                // 添加鼠标事件
                button.MouseEnter += (s, e) => button.FlatAppearance.BorderColor = DefaultTheme.HighlightColor;
                button.MouseLeave += (s, e) => button.FlatAppearance.BorderColor = Color.Transparent;
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化Button失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化TextBox控件
        /// </summary>
        public static void EnhanceTextBox(TextBox textBox, bool useModernStyle = true)
        {
            try
            {
                if (useModernStyle)
                {
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                    textBox.BackColor = DefaultTheme.SurfaceColor;
                    textBox.ForeColor = DefaultTheme.TextColor;
                    textBox.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                    
                    // 添加焦点事件
                    textBox.Enter += (s, e) => 
                    {
                        textBox.BackColor = Color.FromArgb(
                            Math.Min(255, DefaultTheme.SurfaceColor.R + 10),
                            Math.Min(255, DefaultTheme.SurfaceColor.G + 10),
                            Math.Min(255, DefaultTheme.SurfaceColor.B + 10)
                        );
                        textBox.BorderStyle = BorderStyle.FixedSingle;
                    };
                    
                    textBox.Leave += (s, e) => 
                    {
                        textBox.BackColor = DefaultTheme.SurfaceColor;
                        textBox.BorderStyle = BorderStyle.FixedSingle;
                    };
                }
                else
                {
                    textBox.BorderStyle = BorderStyle.Fixed3D;
                    textBox.BackColor = SystemColors.Window;
                    textBox.ForeColor = SystemColors.WindowText;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化TextBox失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化CheckBox控件
        /// </summary>
        public static void EnhanceCheckBox(CheckBox checkBox, bool useModernStyle = true)
        {
            try
            {
                if (useModernStyle)
                {
                    checkBox.FlatStyle = FlatStyle.Flat;
                    checkBox.ForeColor = DefaultTheme.TextColor;
                    checkBox.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                    checkBox.BackColor = Color.Transparent;
                }
                else
                {
                    checkBox.FlatStyle = FlatStyle.Standard;
                    checkBox.ForeColor = SystemColors.ControlText;
                    checkBox.BackColor = SystemColors.Control;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化CheckBox失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化ProgressBar控件
        /// </summary>
        public static void EnhanceProgressBar(ProgressBar progressBar, bool useModernStyle = true)
        {
            try
            {
                if (useModernStyle)
                {
                    progressBar.Style = ProgressBarStyle.Continuous;
                    progressBar.ForeColor = DefaultTheme.PrimaryColor;
                    progressBar.BackColor = DefaultTheme.SurfaceColor;
                }
                else
                {
                    progressBar.Style = ProgressBarStyle.Continuous;
                    progressBar.ForeColor = SystemColors.Highlight;
                    progressBar.BackColor = SystemColors.Control;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化ProgressBar失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化Label控件
        /// </summary>
        public static void EnhanceLabel(Label label, LabelStyle style = LabelStyle.Normal)
        {
            try
            {
                label.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                label.BackColor = Color.Transparent;

                switch (style)
                {
                    case LabelStyle.Normal:
                        label.ForeColor = DefaultTheme.TextColor;
                        break;
                    case LabelStyle.Secondary:
                        label.ForeColor = DefaultTheme.TextSecondaryColor;
                        break;
                    case LabelStyle.Primary:
                        label.ForeColor = DefaultTheme.PrimaryColor;
                        break;
                    case LabelStyle.Success:
                        label.ForeColor = DefaultTheme.SuccessColor;
                        break;
                    case LabelStyle.Warning:
                        label.ForeColor = DefaultTheme.WarningColor;
                        break;
                    case LabelStyle.Error:
                        label.ForeColor = DefaultTheme.ErrorColor;
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化Label失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化GroupBox控件
        /// </summary>
        public static void EnhanceGroupBox(GroupBox groupBox, bool useModernStyle = true)
        {
            try
            {
                if (useModernStyle)
                {
                    groupBox.ForeColor = DefaultTheme.TextColor;
                    groupBox.BackColor = Color.Transparent;
                    groupBox.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                }
                else
                {
                    groupBox.ForeColor = SystemColors.ControlText;
                    groupBox.BackColor = SystemColors.Control;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化GroupBox失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化DataGridView控件
        /// </summary>
        public static void EnhanceDataGridView(DataGridView dataGridView, bool useModernStyle = true)
        {
            try
            {
                if (useModernStyle)
                {
                    dataGridView.BackgroundColor = DefaultTheme.BackgroundColor;
                    dataGridView.GridColor = DefaultTheme.BorderColor;
                    dataGridView.BorderStyle = BorderStyle.FixedSingle;
                    dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
                    dataGridView.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                    dataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                    
                    // 设置字体
                    dataGridView.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                    dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font("微软雅黑", 9F, FontStyle.Bold);
                    
                    // 设置颜色
                    dataGridView.DefaultCellStyle.BackColor = DefaultTheme.SurfaceColor;
                    dataGridView.DefaultCellStyle.ForeColor = DefaultTheme.TextColor;
                    dataGridView.DefaultCellStyle.SelectionBackColor = DefaultTheme.HighlightColor;
                    dataGridView.DefaultCellStyle.SelectionForeColor = Color.White;
                    
                    dataGridView.ColumnHeadersDefaultCellStyle.BackColor = DefaultTheme.SecondaryColor;
                    dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = DefaultTheme.TextColor;
                    dataGridView.ColumnHeadersDefaultCellStyle.SelectionBackColor = DefaultTheme.SecondaryColor;
                    
                    dataGridView.RowHeadersDefaultCellStyle.BackColor = DefaultTheme.SecondaryColor;
                    dataGridView.RowHeadersDefaultCellStyle.ForeColor = DefaultTheme.TextColor;
                    dataGridView.RowHeadersDefaultCellStyle.SelectionBackColor = DefaultTheme.SecondaryColor;
                    
                    // 设置选择模式
                    dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    dataGridView.AllowUserToAddRows = false;
                    dataGridView.AllowUserToDeleteRows = false;
                    dataGridView.ReadOnly = true;
                    dataGridView.MultiSelect = false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化DataGridView失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化整个窗体
        /// </summary>
        public static void EnhanceForm(Form form, bool useModernStyle = true)
        {
            try
            {
                if (useModernStyle)
                {
                    form.BackColor = DefaultTheme.BackgroundColor;
                    form.ForeColor = DefaultTheme.TextColor;
                    form.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                    
                    // 设置窗体样式
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.StartPosition = FormStartPosition.CenterScreen;
                    
                    // 启用双缓冲
                    typeof(Control).InvokeMember("DoubleBuffered", 
                        System.Reflection.BindingFlags.SetProperty | 
                        System.Reflection.BindingFlags.Instance | 
                        System.Reflection.BindingFlags.NonPublic, 
                        null, form, new object[] { true });
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化Form失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 美化所有控件
        /// </summary>
        public static void EnhanceAllControls(Control parent, bool useModernStyle = true)
        {
            try
            {
                foreach (Control control in parent.Controls)
                {
                    // 根据控件类型应用相应的美化
                    if (control is ComboBox comboBox)
                    {
                        EnhanceComboBox(comboBox, useModernStyle);
                    }
                    else if (control is Button button)
                    {
                        EnhanceButton(button, ButtonStyle.Primary);
                    }
                    else if (control is TextBox textBox)
                    {
                        EnhanceTextBox(textBox, useModernStyle);
                    }
                    else if (control is CheckBox checkBox)
                    {
                        EnhanceCheckBox(checkBox, useModernStyle);
                    }
                    else if (control is ProgressBar progressBar)
                    {
                        EnhanceProgressBar(progressBar, useModernStyle);
                    }
                    else if (control is Label label)
                    {
                        EnhanceLabel(label, LabelStyle.Normal);
                    }
                    else if (control is GroupBox groupBox)
                    {
                        EnhanceGroupBox(groupBox, useModernStyle);
                    }
                    else if (control is DataGridView dataGridView)
                    {
                        EnhanceDataGridView(dataGridView, useModernStyle);
                    }

                    // 递归处理子控件
                    if (control.HasChildren)
                    {
                        EnhanceAllControls(control, useModernStyle);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"美化所有控件失败: {ex.Message}", ex);
            }
        }

        #endregion

        #region 样式枚举

        /// <summary>
        /// 按钮样式
        /// </summary>
        public enum ButtonStyle
        {
            Primary,
            Secondary,
            Success,
            Warning,
            Danger
        }

        /// <summary>
        /// 标签样式
        /// </summary>
        public enum LabelStyle
        {
            Normal,
            Secondary,
            Primary,
            Success,
            Warning,
            Error
        }

        #endregion
    }
}