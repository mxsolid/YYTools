using System.Windows;

namespace YYTools.Wpf8
{
    /// <summary>
    /// 应用程序入口，负责全局主题与异常处理初始化
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            // TODO: 可在此加载配置、初始化日志、设置DPI感知等
        }
    }
}

