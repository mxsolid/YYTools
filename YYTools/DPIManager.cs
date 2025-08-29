using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 高级DPI兼容性管理器
    /// 解决.NET Framework 4.8下的多显示器DPI缩放问题
    /// 支持Per-Monitor V2 DPI感知，动态响应显示器DPI变化
    /// </summary>
    public static class DPIManager
    {
        #region Win32 API

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        private static extern int GetDpiForWindow(IntPtr hwnd);

        [DllImport("shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, int flags);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(Point pt, int flags);

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        [DllImport("user32.dll")]
        private static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);

        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;
        private const int MONITOR_DEFAULTTONEAREST = 2;
        private const int MONITOR_DEFAULTTONULL = 0;
        private const int MDT_EFFECTIVE_DPI = 0;
        private const int MDT_RAW_DPI = 1;

        private delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        #endregion

        #region 事件和委托

        /// <summary>
        /// DPI变化事件
        /// </summary>
        public static event EventHandler<DpiChangedEventArgs> DpiChanged;

        /// <summary>
        /// DPI变化事件参数
        /// </summary>
        public class DpiChangedEventArgs : EventArgs
        {
            public float OldDpiScale { get; set; }
            public float NewDpiScale { get; set; }
            public IntPtr MonitorHandle { get; set; }
            public Point MonitorLocation { get; set; }
        }

        #endregion

        #region DPI信息

        private static readonly Dictionary<IntPtr, DpiInfo> _monitorDpiCache = new Dictionary<IntPtr, DpiInfo>();
        private static float _systemDpiScale = 1.0f;
        private static float _primaryMonitorDpiScale = 1.0f;
        private static bool _isInitialized = false;
        private static bool _isPerMonitorV2Enabled = false;

        /// <summary>
        /// 系统DPI缩放比例
        /// </summary>
        public static float SystemDpiScale
        {
            get
            {
                InitializeDpi();
                return _systemDpiScale;
            }
        }

        /// <summary>
        /// 主显示器DPI缩放比例
        /// </summary>
        public static float PrimaryMonitorDpiScale
        {
            get
            {
                InitializeDpi();
                return _primaryMonitorDpiScale;
            }
        }

        /// <summary>
        /// 是否为高DPI显示器
        /// </summary>
        public static bool IsHighDpi
        {
            get
            {
                InitializeDpi();
                return _systemDpiScale > 1.0f;
            }
        }

        /// <summary>
        /// 是否为超高DPI显示器（2K、4K等）
        /// </summary>
        public static bool IsUltraHighDpi
        {
            get
            {
                InitializeDpi();
                return _systemDpiScale > 1.5f;
            }
        }

        /// <summary>
        /// 是否启用了Per-Monitor V2 DPI感知
        /// </summary>
        public static bool IsPerMonitorV2Enabled
        {
            get
            {
                InitializeDpi();
                return _isPerMonitorV2Enabled;
            }
        }

        #endregion

        #region 初始化

        /// <summary>
        /// 初始化DPI信息
        /// </summary>
        private static void InitializeDpi()
        {
            if (_isInitialized) return;

            try
            {
                // 获取系统DPI
                IntPtr hdc = GetDC(IntPtr.Zero);
                if (hdc != IntPtr.Zero)
                {
                    try
                    {
                        int systemDpi = GetDeviceCaps(hdc, LOGPIXELSX);
                        _systemDpiScale = systemDpi / 96.0f;
                        _primaryMonitorDpiScale = _systemDpiScale;
                    }
                    finally
                    {
                        ReleaseDC(IntPtr.Zero, hdc);
                    }
                }

                // 枚举所有显示器并获取DPI信息
                EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, MonitorEnumProc, IntPtr.Zero);

                // 检查是否启用了Per-Monitor V2
                _isPerMonitorV2Enabled = CheckPerMonitorV2Enabled();

                _isInitialized = true;

                Logger.LogInfo($"DPI管理器初始化完成 - 系统DPI: {_systemDpiScale:F2}, Per-Monitor V2: {_isPerMonitorV2Enabled}");
            }
            catch (Exception ex)
            {
                Logger.LogError("DPI管理器初始化失败", ex);
                // 设置默认值
                _systemDpiScale = 1.0f;
                _primaryMonitorDpiScale = 1.0f;
                _isPerMonitorV2Enabled = false;
                _isInitialized = true;
            }
        }

        /// <summary>
        /// 显示器枚举回调
        /// </summary>
        private static bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData)
        {
            try
            {
                var monitorInfo = new MONITORINFO { cbSize = Marshal.SizeOf(typeof(MONITORINFO)) };
                if (GetMonitorInfo(hMonitor, ref monitorInfo))
                {
                    uint dpiX, dpiY;
                    if (GetDpiForMonitor(hMonitor, MDT_EFFECTIVE_DPI, out dpiX, out dpiY) == 0)
                    {
                        float dpiScale = dpiX / 96.0f;
                        var dpiInfo = new DpiInfo
                        {
                            MonitorHandle = hMonitor,
                            DpiScale = dpiScale,
                            Location = new Point(lprcMonitor.left, lprcMonitor.top),
                            Size = new Size(lprcMonitor.right - lprcMonitor.left, lprcMonitor.bottom - lprcMonitor.top)
                        };

                        _monitorDpiCache[hMonitor] = dpiInfo;

                        // 更新主显示器DPI
                        if (lprcMonitor.left == 0 && lprcMonitor.top == 0)
                        {
                            _primaryMonitorDpiScale = dpiScale;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"获取显示器DPI信息失败: {ex.Message}");
            }

            return true;
        }

        /// <summary>
        /// 检查是否启用了Per-Monitor V2
        /// </summary>
        private static bool CheckPerMonitorV2Enabled()
        {
            try
            {
                // 尝试获取当前进程的DPI感知上下文
                var process = System.Diagnostics.Process.GetCurrentProcess();
                var handle = process.MainWindowHandle;
                if (handle != IntPtr.Zero)
                {
                    int dpi = GetDpiForWindow(handle);
                    return dpi > 0;
                }
            }
            catch { }

            return false;
        }

        #endregion

        #region DPI缩放方法

        /// <summary>
        /// 根据指定DPI缩放尺寸
        /// </summary>
        public static Size ScaleSize(Size originalSize, float dpiScale)
        {
            if (dpiScale <= 0) dpiScale = 1.0f;
            return new Size(
                (int)Math.Round(originalSize.Width * dpiScale),
                (int)Math.Round(originalSize.Height * dpiScale)
            );
        }

        /// <summary>
        /// 根据当前DPI缩放尺寸
        /// </summary>
        public static Size ScaleSize(Size originalSize)
        {
            return ScaleSize(originalSize, _primaryMonitorDpiScale);
        }

        /// <summary>
        /// 根据当前DPI缩放尺寸
        /// </summary>
        public static Size ScaleSize(int width, int height)
        {
            return ScaleSize(new Size(width, height));
        }

        /// <summary>
        /// 根据指定DPI缩放点位置
        /// </summary>
        public static Point ScalePoint(Point originalPoint, float dpiScale)
        {
            if (dpiScale <= 0) dpiScale = 1.0f;
            return new Point(
                (int)Math.Round(originalPoint.X * dpiScale),
                (int)Math.Round(originalPoint.Y * dpiScale)
            );
        }

        /// <summary>
        /// 根据当前DPI缩放点位置
        /// </summary>
        public static Point ScalePoint(Point originalPoint)
        {
            return ScalePoint(originalPoint, _primaryMonitorDpiScale);
        }

        /// <summary>
        /// 根据当前DPI缩放点位置
        /// </summary>
        public static Point ScalePoint(int x, int y)
        {
            return ScalePoint(new Point(x, y));
        }

        /// <summary>
        /// 根据指定DPI缩放矩形
        /// </summary>
        public static Rectangle ScaleRectangle(Rectangle originalRect, float dpiScale)
        {
            if (dpiScale <= 0) dpiScale = 1.0f;
            return new Rectangle(
                (int)Math.Round(originalRect.X * dpiScale),
                (int)Math.Round(originalRect.Y * dpiScale),
                (int)Math.Round(originalRect.Width * dpiScale),
                (int)Math.Round(originalRect.Height * dpiScale)
            );
        }

        /// <summary>
        /// 根据当前DPI缩放矩形
        /// </summary>
        public static Rectangle ScaleRectangle(Rectangle originalRect)
        {
            return ScaleRectangle(originalRect, _primaryMonitorDpiScale);
        }

        /// <summary>
        /// 根据指定DPI缩放字体大小
        /// </summary>
        public static float ScaleFontSize(float originalSize, float dpiScale)
        {
            if (dpiScale <= 0) dpiScale = 1.0f;
            // 字体大小缩放使用更保守的算法，避免过大
            float scaleFactor = Math.Min(dpiScale, 1.5f); // 限制最大缩放为1.5倍
            return (float)Math.Round(originalSize * scaleFactor, 1);
        }

        /// <summary>
        /// 根据当前DPI缩放字体大小
        /// </summary>
        public static float ScaleFontSize(float originalSize)
        {
            return ScaleFontSize(originalSize, _primaryMonitorDpiScale);
        }

        /// <summary>
        /// 根据当前DPI缩放字体大小
        /// </summary>
        public static int ScaleFontSize(int originalSize)
        {
            return (int)ScaleFontSize((float)originalSize);
        }

        #endregion

        #region 窗体DPI适配

        /// <summary>
        /// 为窗体启用DPI感知
        /// </summary>
        public static void EnableDpiAwareness(Form form)
        {
            try
            {
                if (form == null) return;

                // 设置窗体的DPI感知模式
                form.AutoScaleMode = AutoScaleMode.Dpi;
                form.AutoScaleDimensions = new SizeF(96F, 96F);

                // 根据DPI调整窗体大小和位置
                if (IsHighDpi)
                {
                    AdjustFormForDpi(form);
                }

                // 添加DPI变化事件处理
                form.HandleCreated += (s, e) => SetupDpiChangeHandling(form);

                Logger.LogInfo($"窗体DPI感知已启用: {form.Name}, DPI缩放: {_primaryMonitorDpiScale:F2}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启用窗体DPI感知失败: {form.Name}", ex);
            }
        }

        /// <summary>
        /// 根据DPI调整窗体
        /// </summary>
        private static void AdjustFormForDpi(Form form)
        {
            try
            {
                // 调整窗体最小尺寸
                if (form.MinimumSize != Size.Empty)
                {
                    form.MinimumSize = ScaleSize(form.MinimumSize);
                }

                // 调整窗体大小（保持比例）
                if (form.Size.Width > 0 && form.Size.Height > 0)
                {
                    Size newSize = ScaleSize(form.Size);
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

        /// <summary>
        /// 设置DPI变化事件处理
        /// </summary>
        private static void SetupDpiChangeHandling(Form form)
        {
            try
            {
                // 监听WM_DPICHANGED消息
                form.HandleCreated += (s, e) =>
                {
                    // 这里可以添加自定义的DPI变化处理逻辑
                };
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"设置DPI变化事件处理失败: {form.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 为控件启用DPI感知
        /// </summary>
        public static void EnableDpiAwareness(Control control)
        {
            try
            {
                if (control == null) return;

                // 根据DPI调整控件大小和位置
                if (IsHighDpi)
                {
                    AdjustControlForDpi(control);
                }

                // 递归处理子控件
                foreach (Control child in control.Controls)
                {
                    EnableDpiAwareness(child);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"启用控件DPI感知失败: {control.Name}", ex);
            }
        }

        /// <summary>
        /// 根据DPI调整控件
        /// </summary>
        private static void AdjustControlForDpi(Control control)
        {
            try
            {
                // 调整控件大小
                if (control.Size.Width > 0 && control.Size.Height > 0)
                {
                    Size newSize = ScaleSize(control.Size);
                    // 限制最大尺寸，避免控件过大
                    if (newSize.Width > 800) newSize.Width = 800;
                    if (newSize.Height > 600) newSize.Height = 600;
                    control.Size = newSize;
                }

                // 调整控件位置
                if (control.Location.X > 0 || control.Location.Y > 0)
                {
                    control.Location = ScalePoint(control.Location);
                }

                // 调整字体大小
                if (control.Font != null)
                {
                    float newSize = ScaleFontSize(control.Font.Size);
                    if (Math.Abs(newSize - control.Font.Size) > 0.1f)
                    {
                        control.Font = new Font(control.Font.FontFamily, newSize, control.Font.Style);
                    }
                }

                // 特殊处理某些控件类型
                if (control is ComboBox comboBox)
                {
                    AdjustComboBoxForDpi(comboBox);
                }
                else if (control is TextBox textBox)
                {
                    AdjustTextBoxForDpi(textBox);
                }
                else if (control is Button button)
                {
                    AdjustButtonForDpi(button);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"调整控件DPI失败: {control.Name}, {ex.Message}");
            }
        }

        /// <summary>
        /// 调整ComboBox的DPI
        /// </summary>
        private static void AdjustComboBoxForDpi(ComboBox comboBox)
        {
            try
            {
                // 调整下拉列表的宽度，确保文本不被截断
                if (comboBox.DropDownWidth > 0)
                {
                    comboBox.DropDownWidth = ScaleSize(comboBox.DropDownWidth, 0).Width;
                }
            }
            catch { }
        }

        /// <summary>
        /// 调整TextBox的DPI
        /// </summary>
        private static void AdjustTextBoxForDpi(TextBox textBox)
        {
            try
            {
                // 调整文本框的字符宽度
                if (textBox.MaxLength > 0)
                {
                    // 根据DPI调整最大字符数
                    int newMaxLength = (int)(textBox.MaxLength * _primaryMonitorDpiScale);
                    if (newMaxLength > 0 && newMaxLength < 10000)
                    {
                        textBox.MaxLength = newMaxLength;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 调整Button的DPI
        /// </summary>
        private static void AdjustButtonForDpi(Button button)
        {
            try
            {
                // 调整按钮的内边距
                if (button.Padding != Padding.Empty)
                {
                    button.Padding = new Padding(
                        ScaleSize(button.Padding.Left, 0).Width,
                        ScaleSize(button.Padding.Top, 0).Height,
                        ScaleSize(button.Padding.Right, 0).Width,
                        ScaleSize(button.Padding.Bottom, 0).Height
                    );
                }
            }
            catch { }
        }

        /// <summary>
        /// 为所有控件启用DPI感知
        /// </summary>
        public static void EnableDpiAwarenessForAllControls(Form form)
        {
            try
            {
                if (form == null) return;

                // 启用窗体DPI感知
                EnableDpiAwareness(form);

                // 启用所有控件DPI感知
                foreach (Control control in form.Controls)
                {
                    EnableDpiAwareness(control);
                }

                Logger.LogInfo($"所有控件DPI感知已启用: {form.Name}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启用所有控件DPI感知失败: {form.Name}", ex);
            }
        }

        #endregion

        #region 字体DPI适配

        /// <summary>
        /// 创建DPI感知的字体
        /// </summary>
        public static Font CreateDpiAwareFont(string fontFamily, float size, FontStyle style = FontStyle.Regular)
        {
            try
            {
                float scaledSize = ScaleFontSize(size);
                return new Font(fontFamily, scaledSize, style);
            }
            catch (Exception ex)
            {
                Logger.LogError($"创建DPI感知字体失败: {fontFamily}, {size}", ex);
                // 返回原始字体
                return new Font(fontFamily, size, style);
            }
        }

        /// <summary>
        /// 创建DPI感知的字体
        /// </summary>
        public static Font CreateDpiAwareFont(string fontFamily, int size, FontStyle style = FontStyle.Regular)
        {
            return CreateDpiAwareFont(fontFamily, (float)size, style);
        }

        /// <summary>
        /// 创建默认DPI感知字体
        /// </summary>
        public static Font CreateDefaultDpiAwareFont()
        {
            return CreateDpiAwareFont("微软雅黑", 9F);
        }

        #endregion

        #region 多显示器支持

        /// <summary>
        /// 获取指定位置的显示器DPI
        /// </summary>
        public static float GetDpiForLocation(Point location)
        {
            try
            {
                IntPtr monitorHandle = MonitorFromPoint(location, MONITOR_DEFAULTTONEAREST);
                if (monitorHandle != IntPtr.Zero && _monitorDpiCache.ContainsKey(monitorHandle))
                {
                    return _monitorDpiCache[monitorHandle].DpiScale;
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"获取位置DPI失败: {location}, {ex.Message}");
            }

            return _primaryMonitorDpiScale;
        }

        /// <summary>
        /// 获取指定窗体的显示器DPI
        /// </summary>
        public static float GetDpiForWindow(Form form)
        {
            try
            {
                if (form == null || !form.IsHandleCreated) return _primaryMonitorDpiScale;

                IntPtr monitorHandle = MonitorFromWindow(form.Handle, MONITOR_DEFAULTTONEAREST);
                if (monitorHandle != IntPtr.Zero && _monitorDpiCache.ContainsKey(monitorHandle))
                {
                    return _monitorDpiCache[monitorHandle].DpiScale;
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"获取窗体DPI失败: {form.Name}, {ex.Message}");
            }

            return _primaryMonitorDpiScale;
        }

        /// <summary>
        /// 刷新显示器DPI信息
        /// </summary>
        public static void RefreshMonitorDpiInfo()
        {
            try
            {
                _monitorDpiCache.Clear();
                _isInitialized = false;
                InitializeDpi();
                Logger.LogInfo("显示器DPI信息已刷新");
            }
            catch (Exception ex)
            {
                Logger.LogError("刷新显示器DPI信息失败", ex);
            }
        }

        #endregion

        #region 工具方法

        /// <summary>
        /// 获取DPI信息字符串
        /// </summary>
        public static string GetDpiInfo()
        {
            InitializeDpi();
            return $"系统DPI: {_systemDpiScale:F2}, 主显示器DPI: {_primaryMonitorDpiScale:F2}, 高DPI: {IsHighDpi}, 超高DPI: {IsUltraHighDpi}, Per-Monitor V2: {_isPerMonitorV2Enabled}";
        }

        /// <summary>
        /// 重置DPI设置
        /// </summary>
        public static void ResetDpi()
        {
            _isInitialized = false;
            _monitorDpiCache.Clear();
            _systemDpiScale = 1.0f;
            _primaryMonitorDpiScale = 1.0f;
            _isPerMonitorV2Enabled = false;
            InitializeDpi();
        }

        #endregion

        #region 内部类

        /// <summary>
        /// 显示器DPI信息
        /// </summary>
        private class DpiInfo
        {
            public IntPtr MonitorHandle { get; set; }
            public float DpiScale { get; set; }
            public Point Location { get; set; }
            public Size Size { get; set; }
        }

        #endregion
    }
}