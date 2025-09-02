using System;
using System.Windows;

namespace YYTools.Wpf8
{
    /// <summary>
    /// WPF程序入口 (保留以便 Rider 启动识别)
    /// </summary>
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            var app = new App();
            app.InitializeComponent();
            app.Run();
        }
    }
}

