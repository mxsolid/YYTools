using System.Windows;
using MahApps.Metro.Controls;
using MahApps.Metro.Theming;

namespace YYTools.Wpf8.Views
{
    /// <summary>
    /// 主窗口，集成MahApps与HandyControl，并支持主题切换
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private bool _isDark = false;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnToggleThemeClick(object sender, RoutedEventArgs e)
        {
            _isDark = !_isDark;
            var theme = _isDark ? ThemeManager.BaseColorDark : ThemeManager.BaseColorLight;
            ThemeManager.Current.ChangeThemeBaseColor(Application.Current, theme);
        }
    }
}

