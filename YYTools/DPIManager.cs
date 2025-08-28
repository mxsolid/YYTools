using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// DPI兼容性管理器
    /// 解决.NET Framework 4.8下的DPI缩放问题
    /// </summary>
    public static class DPIManager
    {
        #region Win32 API

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        private static extern int GetDpiForWindow(IntPtr hwnd);

        [DllImport("shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, int flags);

        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;
        private const int MONITOR_DEFAULTTONEAREST = 2;
        private const int MDT_EFFECTIVE_DPI = 0;

        #endregion

        #region DPI信息

        private static float? _systemDpiX;
        private static float? _systemDpiY;
        private static float? _currentDpiX;
        private static float? _currentDpiY;
        private static bool _isInitialized = false;

        /// <summary>
        /// 系统DPI缩放比例
        /// </summary>
        public static float SystemDpiScale
        {
            get
            {
                InitializeDpi();
                return _systemDpiX ?? 1.0f;
            }
        }

        /// <summary>
        /// 当前DPI缩放比例
        /// </summary>
        public static float CurrentDpiScale
        {
            get
            {
                InitializeDpi();
                return _currentDpiX ?? 1.0f;
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
                return (_systemDpiX ?? 96) > 96;
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
                return (_systemDpiX ?? 96) > 144;
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
                // 使用更兼容的方法获取DPI信息
                _systemDpiX = 1.0f;
                _systemDpiY = 1.0f;
                
                // 尝试使用Graphics获取DPI信息
                using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
                {
                    try
                    {
                        _systemDpiX = graphics.DpiX / 96.0f;
                        _systemDpiY = graphics.DpiY / 96.0f;
                    }
                    catch
                    {
                        // 如果Graphics方法失败，使用默认值
                        _systemDpiX = 1.0f;
                        _systemDpiY = 1.0f;
                    }
                }

                // 设置当前DPI为系统DPI
                _currentDpiX = _systemDpiX;
                _currentDpiY = _systemDpiY;

                _isInitialized = true;

                Logger.LogInfo($"DPI管理器初始化完成 - 系统DPI: {_systemDpiX:F2}x{_systemDpiY:F2}");
            }
            catch (Exception ex)
            {
                Logger.LogError("DPI管理器初始化失败", ex);
                // 设置默认值
                _systemDpiX = 1.0f;
                _systemDpiY = 1.0f;
                _currentDpiX = 1.0f;
                _currentDpiY = 1.0f;
                _isInitialized = true;
            }
        }

        #endregion

        #region DPI缩放方法

        /// <summary>
        /// 根据DPI缩放尺寸
        /// </summary>
        public static Size ScaleSize(Size originalSize)
        {
            float scale = CurrentDpiScale;
            return new Size(
                (int)(originalSize.Width * scale),
                (int)(originalSize.Height * scale)
            );
        }

        /// <summary>
        /// 根据DPI缩放尺寸
        /// </summary>
        public static Size ScaleSize(int width, int height)
        {
            return ScaleSize(new Size(width, height));
        }

        /// <summary>
        /// 根据DPI缩放点位置
        /// </summary>
        public static Point ScalePoint(Point originalPoint)
        {
            float scale = CurrentDpiScale;
            return new Point(
                (int)(originalPoint.X * scale),
                (int)(originalPoint.Y * scale)
            );
        }

        /// <summary>
        /// 根据DPI缩放点位置
        /// </summary>
        public static Point ScalePoint(int x, int y)
        {
            return ScalePoint(new Point(x, y));
        }

        /// <summary>
        /// 根据DPI缩放矩形
        /// </summary>
        public static Rectangle ScaleRectangle(Rectangle originalRect)
        {
            float scale = CurrentDpiScale;
            return new Rectangle(
                (int)(originalRect.X * scale),
                (int)(originalRect.Y * scale),
                (int)(originalRect.Width * scale),
                (int)(originalRect.Height * scale)
            );
        }

        /// <summary>
        /// 根据DPI缩放字体大小
        /// </summary>
        public static float ScaleFontSize(float originalSize)
        {
            float scale = CurrentDpiScale;
            return originalSize * scale;
        }

        /// <summary>
        /// 根据DPI缩放字体大小
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

                // 设置窗体的DPI感知
                form.AutoScaleMode = AutoScaleMode.Dpi;
                form.AutoScaleDimensions = new SizeF(96F, 96F);

                // 根据DPI调整窗体大小
                if (IsHighDpi)
                {
                    // 调整窗体最小尺寸
                    if (form.MinimumSize != Size.Empty)
                    {
                        form.MinimumSize = ScaleSize(form.MinimumSize);
                    }

                    // 调整窗体大小
                    if (form.Size.Width > 0 && form.Size.Height > 0)
                    {
                        form.Size = ScaleSize(form.Size);
                    }

                    // 调整窗体位置
                    if (form.StartPosition == FormStartPosition.CenterScreen)
                    {
                        form.StartPosition = FormStartPosition.CenterScreen;
                    }
                }

                Logger.LogInfo($"窗体DPI感知已启用: {form.Name}, DPI缩放: {CurrentDpiScale:F2}");
            }
            catch (Exception ex)
            {
                Logger.LogError($"启用窗体DPI感知失败: {form.Name}", ex);
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
                    // 调整控件大小
                    if (control.Size.Width > 0 && control.Size.Height > 0)
                    {
                        control.Size = ScaleSize(control.Size);
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
                        if (newSize != control.Font.Size)
                        {
                            control.Font = new Font(control.Font.FontFamily, newSize, control.Font.Style);
                        }
                    }

                    // 递归处理子控件
                    foreach (Control child in control.Controls)
                    {
                        EnableDpiAwareness(child);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"启用控件DPI感知失败: {control.Name}", ex);
            }
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

        #region 工具方法

        /// <summary>
        /// 获取DPI信息字符串
        /// </summary>
        public static string GetDpiInfo()
        {
            InitializeDpi();
            return $"系统DPI: {_systemDpiX:F2}x{_systemDpiY:F2}, 当前DPI: {_currentDpiX:F2}x{_currentDpiY:F2}, 高DPI: {IsHighDpi}, 超高DPI: {IsUltraHighDpi}";
        }

        /// <summary>
        /// 重置DPI设置
        /// </summary>
        public static void ResetDpi()
        {
            _isInitialized = false;
            _systemDpiX = null;
            _systemDpiY = null;
            _currentDpiX = null;
            _currentDpiY = null;
            InitializeDpi();
        }

        #endregion
    }
}