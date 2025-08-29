using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.ComponentModel;

namespace YYTools
{
    /// <summary>
    /// UI增强器 - 提供DPI感知的界面优化功能
    /// 解决高DPI显示器下的界面显示问题
    /// </summary>
    public static class UIEnhancer
    {
        #region 常量定义

        /// <summary>
        /// 最大字体缩放比例
        /// </summary>
        private const float MAX_FONT_SCALE = 1.5f;
        
        /// <summary>
        /// 最小字体缩放比例
        /// </summary>
        private const float MIN_FONT_SCALE = 0.8f;
        
        /// <summary>
        /// 最大控件宽度
        /// </summary>
        private const int MAX_CONTROL_WIDTH = 800;
        
        /// <summary>
        /// 最大控件高度
        /// </summary>
        private const int MAX_CONTROL_HEIGHT = 600;

        #endregion

        #region 颜色主题

        /// <summary>
        /// 默认颜色主题
        /// </summary>
        public static class DefaultTheme
        {
            // 切换为浅色主题以统一界面风格
            public static Color PrimaryColor = LightTheme.PrimaryColor;
            public static Color SecondaryColor = LightTheme.SecondaryColor;
            public static Color AccentColor = LightTheme.AccentColor;
            public static Color SuccessColor = LightTheme.SuccessColor;
            public static Color WarningColor = LightTheme.WarningColor;
            public static Color ErrorColor = LightTheme.ErrorColor;
            public static Color BackgroundColor = LightTheme.BackgroundColor;
            public static Color SurfaceColor = LightTheme.SurfaceColor;
            public static Color TextColor = LightTheme.TextColor;
            public static Color TextSecondaryColor = LightTheme.TextSecondaryColor;
            public static Color BorderColor = LightTheme.BorderColor;
            public static Color HighlightColor = LightTheme.HighlightColor;
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

        #region DPI感知的UI优化

        /// <summary>
        /// 为窗体启用完整的DPI感知优化
        /// </summary>
        public static void EnableDpiOptimization(Form form)
        {
            try
            {
                if (form == null) return;

                Logger.LogInfo($"开始为窗体 {form.Name} 启用DPI优化");

                // 设置窗体的DPI感知模式
                form.AutoScaleMode = AutoScaleMode.Dpi;
                form.AutoScaleDimensions = new SizeF(96F, 96F);

                // 根据DPI调整窗体
                AdjustFormForDpi(form);

                // 为所有控件启用DPI优化
                EnableDpiOptimizationForAllControls(form);

                // 添加DPI变化事件处理
                SetupDpiChangeHandling(form);

                Logger.LogInfo($"窗体 {form.Name} DPI优化完成");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启用窗体DPI优化失败: {form.Name}", ex);
            }
        }

        /// <summary>
        /// 为所有控件启用DPI优化
        /// </summary>
        public static void EnableDpiOptimizationForAllControls(Form form)
        {
            try
            {
                if (form == null) return;

                // 递归处理所有控件
                ProcessControlRecursively(form);

                Logger.LogInfo($"所有控件DPI优化完成: {form.Name}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启用所有控件DPI优化失败: {form.Name}", ex);
            }
        }

        /// <summary>
        /// 递归处理控件
        /// </summary>
        private static void ProcessControlRecursively(Control control)
        {
            try
            {
                if (control == null) return;

                // 为当前控件启用DPI优化
                EnableDpiOptimizationForControl(control);

                // 递归处理子控件
                foreach (Control child in control.Controls)
                {
                    ProcessControlRecursively(child);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"处理控件DPI优化失败: {control.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 为单个控件启用DPI优化
        /// </summary>
        private static void EnableDpiOptimizationForControl(Control control)
        {
            try
            {
                if (control == null) return;

                // 根据控件类型进行特殊处理
                if (control is ComboBox comboBox)
                {
                    OptimizeComboBoxForDpi(comboBox);
                }
                else if (control is TextBox textBox)
                {
                    OptimizeTextBoxForDpi(textBox);
                }
                else if (control is Button button)
                {
                    OptimizeButtonForDpi(button);
                }
                else if (control is Label label)
                {
                    OptimizeLabelForDpi(label);
                }
                else if (control is GroupBox groupBox)
                {
                    OptimizeGroupBoxForDpi(groupBox);
                }
                else if (control is Panel panel)
                {
                    OptimizePanelForDpi(panel);
                }
                else
                {
                    // 通用控件优化
                    OptimizeGenericControlForDpi(control);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化控件DPI失败: {control.Name}, {ex.Message}");
            }
        }

        #endregion

        #region 特定控件类型的DPI优化

        /// <summary>
        /// 优化ComboBox的DPI显示
        /// </summary>
        private static void OptimizeComboBoxForDpi(ComboBox comboBox)
        {
            try
            {
                // 调整字体大小
                AdjustFontForDpi(comboBox);

                // 调整控件大小
                AdjustSizeForDpi(comboBox);

                // 调整下拉列表宽度，确保文本不被截断
                if (comboBox.DropDownWidth > 0)
                {
                    int newWidth = (int)(comboBox.DropDownWidth * DPIManager.PrimaryMonitorDpiScale);
                    newWidth = Math.Min(newWidth, MAX_CONTROL_WIDTH);
                    if (newWidth > comboBox.Width)
                    {
                        comboBox.DropDownWidth = newWidth;
                    }
                }

                // 设置合适的项目高度
                if (DPIManager.IsHighDpi)
                {
                    comboBox.ItemHeight = (int)(comboBox.ItemHeight * DPIManager.PrimaryMonitorDpiScale);
                    comboBox.ItemHeight = Math.Min(comboBox.ItemHeight, 50);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化ComboBox DPI失败: {comboBox.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 优化TextBox的DPI显示
        /// </summary>
        private static void OptimizeTextBoxForDpi(TextBox textBox)
        {
            try
            {
                // 调整字体大小
                AdjustFontForDpi(textBox);

                // 调整控件大小
                AdjustSizeForDpi(textBox);

                // 调整字符宽度相关的设置
                if (textBox.MaxLength > 0 && DPIManager.IsHighDpi)
                {
                    // 根据DPI调整最大字符数，但保持合理范围
                    int newMaxLength = (int)(textBox.MaxLength * DPIManager.PrimaryMonitorDpiScale);
                    if (newMaxLength > 0 && newMaxLength < 10000)
                    {
                        textBox.MaxLength = newMaxLength;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化TextBox DPI失败: {textBox.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 优化Button的DPI显示
        /// </summary>
        private static void OptimizeButtonForDpi(Button button)
        {
            try
            {
                // 调整字体大小
                AdjustFontForDpi(button);

                // 调整控件大小
                AdjustSizeForDpi(button);

                // 调整按钮的内边距
                if (button.Padding != Padding.Empty)
                {
                    button.Padding = ScalePadding(button.Padding);
                }

                // 调整按钮的边距
                if (button.Margin != Padding.Empty)
                {
                    button.Margin = ScalePadding(button.Margin);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化Button DPI失败: {button.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 优化Label的DPI显示
        /// </summary>
        private static void OptimizeLabelForDpi(Label label)
        {
            try
            {
                // 调整字体大小
                AdjustFontForDpi(label);

                // 调整控件大小
                AdjustSizeForDpi(label);

                // 确保标签文本能够完整显示
                if (label.AutoSize && !string.IsNullOrEmpty(label.Text))
                {
                    // 计算文本所需的最小尺寸
                    using (Graphics g = label.CreateGraphics())
                    {
                        SizeF textSize = g.MeasureString(label.Text, label.Font);
                        Size newSize = new Size(
                            (int)Math.Ceiling(textSize.Width),
                            (int)Math.Ceiling(textSize.Height)
                        );
                        
                        // 限制最大尺寸
                        newSize.Width = Math.Min(newSize.Width, MAX_CONTROL_WIDTH);
                        newSize.Height = Math.Min(newSize.Height, MAX_CONTROL_HEIGHT);
                        
                        label.Size = newSize;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化Label DPI失败: {label.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 优化GroupBox的DPI显示
        /// </summary>
        private static void OptimizeGroupBoxForDpi(GroupBox groupBox)
        {
            try
            {
                // 调整字体大小
                AdjustFontForDpi(groupBox);

                // 调整控件大小
                AdjustSizeForDpi(groupBox);

                // 调整GroupBox的标题位置
                if (DPIManager.IsHighDpi)
                {
                    // 确保标题文本不被截断
                    using (Graphics g = groupBox.CreateGraphics())
                    {
                        SizeF textSize = g.MeasureString(groupBox.Text, groupBox.Font);
                        int minWidth = (int)Math.Ceiling(textSize.Width) + 20;
                        if (groupBox.Width < minWidth)
                        {
                            groupBox.Width = Math.Min(minWidth, MAX_CONTROL_WIDTH);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化GroupBox DPI失败: {groupBox.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 优化Panel的DPI显示
        /// </summary>
        private static void OptimizePanelForDpi(Panel panel)
        {
            try
            {
                // 调整控件大小
                AdjustSizeForDpi(panel);

                // 调整Panel的内边距
                if (panel.Padding != Padding.Empty)
                {
                    panel.Padding = ScalePadding(panel.Padding);
                }

                // 调整Panel的边距
                if (panel.Margin != Padding.Empty)
                {
                    panel.Margin = ScalePadding(panel.Margin);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化Panel DPI失败: {panel.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 通用控件的DPI优化
        /// </summary>
        private static void OptimizeGenericControlForDpi(Control control)
        {
            try
            {
                // 调整字体大小
                AdjustFontForDpi(control);

                // 调整控件大小
                AdjustSizeForDpi(control);

                // 调整控件位置
                AdjustLocationForDpi(control);
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"优化通用控件DPI失败: {control.Name}, {ex.Message}");
            }
        }

        #endregion

        #region 通用DPI调整方法

        /// <summary>
        /// 调整字体大小以适应DPI
        /// </summary>
        private static void AdjustFontForDpi(Control control)
        {
            try
            {
                if (control.Font == null) return;

                float newSize = DPIManager.ScaleFontSize(control.Font.Size);
                
                // 限制字体大小范围
                newSize = Math.Max(newSize, control.Font.Size * MIN_FONT_SCALE);
                newSize = Math.Min(newSize, control.Font.Size * MAX_FONT_SCALE);

                if (Math.Abs(newSize - control.Font.Size) > 0.1f)
                {
                    control.Font = new Font(control.Font.FontFamily, newSize, control.Font.Style);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"调整字体DPI失败: {control.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 调整控件大小以适应DPI
        /// </summary>
        private static void AdjustSizeForDpi(Control control)
        {
            try
            {
                if (control.Size.Width <= 0 || control.Size.Height <= 0) return;

                Size newSize = DPIManager.ScaleSize(control.Size);
                
                // 限制最大尺寸
                newSize.Width = Math.Min(newSize.Width, MAX_CONTROL_WIDTH);
                newSize.Height = Math.Min(newSize.Height, MAX_CONTROL_HEIGHT);

                control.Size = newSize;
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"调整控件大小DPI失败: {control.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 调整控件位置以适应DPI
        /// </summary>
        private static void AdjustLocationForDpi(Control control)
        {
            try
            {
                if (control.Location.X < 0 && control.Location.Y < 0) return;

                Point newLocation = DPIManager.ScalePoint(control.Location);
                
                // 确保位置不为负数
                newLocation.X = Math.Max(0, newLocation.X);
                newLocation.Y = Math.Max(0, newLocation.Y);

                control.Location = newLocation;
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"调整控件位置DPI失败: {control.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 缩放Padding
        /// </summary>
        private static Padding ScalePadding(Padding padding)
        {
            try
            {
                if (padding == Padding.Empty) return padding;

                return new Padding(
                    (int)(padding.Left * DPIManager.PrimaryMonitorDpiScale),
                    (int)(padding.Top * DPIManager.PrimaryMonitorDpiScale),
                    (int)(padding.Right * DPIManager.PrimaryMonitorDpiScale),
                    (int)(padding.Bottom * DPIManager.PrimaryMonitorDpiScale)
                );
            }
            catch
            {
                return padding;
            }
        }

        #endregion

        #region 窗体DPI调整

        /// <summary>
        /// 根据DPI调整窗体
        /// </summary>
        private static void AdjustFormForDpi(Form form)
        {
            try
            {
                if (form == null) return;

                // 调整窗体最小尺寸
                if (form.MinimumSize != Size.Empty)
                {
                    Size newMinSize = DPIManager.ScaleSize(form.MinimumSize);
                    newMinSize.Width = Math.Min(newMinSize.Width, MAX_CONTROL_WIDTH);
                    newMinSize.Height = Math.Min(newMinSize.Height, MAX_CONTROL_HEIGHT);
                    form.MinimumSize = newMinSize;
                }

                // 调整窗体大小
                if (form.Size.Width > 0 && form.Size.Height > 0)
                {
                    Size newSize = DPIManager.ScaleSize(form.Size);
                    
                    // 确保窗体不会超出屏幕边界
                    Rectangle screenBounds = Screen.GetWorkingArea(form);
                    if (newSize.Width > screenBounds.Width)
                    {
                        newSize.Width = screenBounds.Width - 100;
                    }
                    if (newSize.Height > screenBounds.Height)
                    {
                        newSize.Height = screenBounds.Height - 100;
                    }
                    
                    // 限制最大尺寸
                    newSize.Width = Math.Min(newSize.Width, MAX_CONTROL_WIDTH);
                    newSize.Height = Math.Min(newSize.Height, MAX_CONTROL_HEIGHT);
                    
                    form.Size = newSize;
                }

                // 调整窗体位置
                if (form.StartPosition == FormStartPosition.CenterScreen)
                {
                    form.StartPosition = FormStartPosition.CenterScreen;
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"调整窗体DPI失败: {form.Name}, {ex.Message}");
            }
        }

        #endregion

        #region DPI变化事件处理

        /// <summary>
        /// 设置DPI变化事件处理
        /// </summary>
        private static void SetupDpiChangeHandling(Form form)
        {
            try
            {
                // 监听DPI变化事件
                DPIManager.DpiChanged += (sender, e) => OnDpiChanged(form, e);
                
                Logger.LogInfo($"窗体 {form.Name} DPI变化事件处理已设置");
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"设置DPI变化事件处理失败: {form.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// DPI变化事件处理
        /// </summary>
        private static void OnDpiChanged(Form form, DPIManager.DpiChangedEventArgs e)
        {
            try
            {
                Logger.LogInfo($"检测到DPI变化: {e.OldDpiScale:F2} -> {e.NewDpiScale:F2}");

                // 在UI线程中处理DPI变化
                if (form.InvokeRequired)
                {
                    form.BeginInvoke(new Action(() => OnDpiChanged(form, e)));
                    return;
                }

                // 刷新显示器DPI信息
                DPIManager.RefreshMonitorDpiInfo();

                // 重新调整界面布局
                AdjustFormForDpi(form);
                EnableDpiOptimizationForAllControls(form);

                // 强制重绘
                form.Invalidate();
                form.Update();

                Logger.LogInfo($"窗体 {form.Name} DPI变化处理完成");
            }
            catch (Exception ex)
            {
                Logger.LogError($"处理DPI变化事件失败: {form.Name}", ex);
            }
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
                    
                    // 在.NET 4.8中ComboBox没有BorderStyle属性，使用其他方式美化
                    comboBox.Font = new Font("微软雅黑", 9F, FontStyle.Regular);
                    comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                    
                    // 添加自定义绘制事件来美化边框
                    comboBox.DrawMode = DrawMode.OwnerDrawFixed;
                    comboBox.DrawItem += (s, e) =>
                    {
                        if (e.Index >= 0)
                        {
                            e.DrawBackground();
                            if (comboBox.Items[e.Index] != null)
                            {
                                TextRenderer.DrawText(e.Graphics, comboBox.Items[e.Index].ToString(),
                                    comboBox.Font, e.Bounds, e.ForeColor, TextFormatFlags.Left);
                            }
                        }
                    };
                }
                else
                {
                    comboBox.FlatStyle = FlatStyle.Standard;
                    comboBox.BackColor = SystemColors.Window;
                    comboBox.ForeColor = SystemColors.WindowText;
                    comboBox.DrawMode = DrawMode.Normal;
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

                // 添加鼠标事件（避免设置为 Transparent 导致 NotSupportedException）
                button.MouseEnter += (s, e) =>
                {
                    try { button.FlatAppearance.BorderSize = 1; button.FlatAppearance.BorderColor = DefaultTheme.HighlightColor; } catch { }
                };
                button.MouseLeave += (s, e) =>
                {
                    try { button.FlatAppearance.BorderSize = 0; /* 不设置为 Transparent，保持当前色 */ } catch { }
                };
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
                    
                    // 设置窗体样式 - 允许调整大小
                    form.FormBorderStyle = FormBorderStyle.Sizable;
                    form.MaximizeBox = true;
                    form.MinimizeBox = true;
                    form.StartPosition = FormStartPosition.CenterScreen;
                    
                    // 启用双缓冲
                    typeof(Control).InvokeMember("DoubleBuffered", 
                        System.Reflection.BindingFlags.SetProperty | 
                        System.Reflection.BindingFlags.Instance | 
                        System.Reflection.BindingFlags.NonPublic, 
                        null, form, new object[] { true });
                        
                    // 设置最小尺寸
                    form.MinimumSize = new Size(800, 600);
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

        #region 工具方法

        /// <summary>
        /// 获取DPI信息字符串
        /// </summary>
        public static string GetDpiInfo()
        {
            return DPIManager.GetDpiInfo();
        }

        /// <summary>
        /// 检查是否为高DPI显示器
        /// </summary>
        public static bool IsHighDpi
        {
            get { return DPIManager.IsHighDpi; }
        }

        /// <summary>
        /// 获取当前DPI缩放比例
        /// </summary>
        public static float CurrentDpiScale
        {
            get { return DPIManager.PrimaryMonitorDpiScale; }
        }

        #endregion
    }
}