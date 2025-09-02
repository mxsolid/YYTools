using System;
using System.IO;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace YYTools.App
{
	/// <summary>
	/// 应用程序入口，负责初始化依赖注入和日志系统。
	/// </summary>
	public partial class App : Application
	{
		public static IServiceProvider Services { get; private set; } = null!;

		protected override void OnStartup(StartupEventArgs e)
		{
			// 初始化 Serilog 日志系统
			var logDir = Path.Combine(AppContext.BaseDirectory, "Logs");
			Directory.CreateDirectory(logDir);
			Log.Logger = new LoggerConfiguration()
				.MinimumLevel.Debug()
				.Enrich.FromLogContext()
				.WriteTo.Console()
				.WriteTo.File(Path.Combine(logDir, $"{DateTime.Now:yyyyMMdd}info.log"),
					rollingInterval: RollingInterval.Day,
					retainedFileCountLimit: 14,
					encoding: System.Text.Encoding.UTF8)
				.CreateLogger();

			Log.Information("系统 - 初始化应用程序...");
			Log.Information("系统 - 日志文件位于: {Path}", logDir);

			var services = new ServiceCollection();
			ConfigureServices(services);
			Services = services.BuildServiceProvider();

			base.OnStartup(e);

			var mainWindow = Services.GetRequiredService<MainWindow>();
			mainWindow.Show();
			Log.Information("系统 - 应用程序启动");
		}

		private static void ConfigureServices(IServiceCollection services)
		{
			// 注册 View 和 ViewModel
			services.AddSingleton<MainWindow>();
			services.AddSingleton<ViewModels.MainViewModel>();

			// 注册核心与服务层
			services.AddSingleton<YYTools.Services.ExcelInteropService>();
			services.AddSingleton<YYTools.Services.ExcelFileParserService>();
			services.AddSingleton<YYTools.Services.MatchServiceV2>();
		}

		protected override void OnExit(ExitEventArgs e)
		{
			Log.Information("系统 - 应用程序退出");
			Log.CloseAndFlush();
			base.OnExit(e);
		}
	}
}