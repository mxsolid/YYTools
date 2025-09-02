using System.Windows;
using HandyControl.Tools;

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
            // 初始化HandyControl语言与主题（中文）
            ConfigHelper.Instance.SetLang("zh-cn");
        }
    }
}

