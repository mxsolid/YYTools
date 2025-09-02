using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Serilog;
using YYTools.Services;

namespace YYTools.App.ViewModels
{
	/// <summary>
	/// 主视图模型：负责UI状态管理、异步任务调度与进度上报。
	/// </summary>
	public partial class MainViewModel : ObservableObject
	{
		private readonly ExcelInteropService _excelInteropService;
		private readonly ExcelFileParserService _excelFileParserService;
		private readonly MatchServiceV2 _matchService;

		public ObservableCollection<string> Logs { get; } = new();
		public ObservableCollection<string> Sheets { get; } = new();
		public ObservableCollection<string> Columns { get; } = new();

		[ObservableProperty]
		private string? selectedSheet;

		[ObservableProperty]
		private string? selectedColumn;

		[ObservableProperty]
		private int progressValue;

		[ObservableProperty]
		private string progressText = "就绪";

		[ObservableProperty]
		private bool isDarkTheme;

		[ObservableProperty]
		private string dpiInfo = "100%";

		public IRelayCommand ReadFromActiveExcelCommand { get; }
		public IRelayCommand PickFileCommand { get; }
		public IRelayCommand StartProcessCommand { get; }

		public MainViewModel(ExcelInteropService excelInteropService, ExcelFileParserService excelFileParserService, MatchServiceV2 matchService)
		{
			_excelInteropService = excelInteropService;
			_excelFileParserService = excelFileParserService;
			_matchService = matchService;

			ReadFromActiveExcelCommand = new AsyncRelayCommand(ReadFromActiveExcelAsync);
			PickFileCommand = new AsyncRelayCommand(PickFileAsync);
			StartProcessCommand = new AsyncRelayCommand(StartProcessAsync);

			Log.Information("[ExcelMerger] ViewModel已构造完成，等待View加载...");
		}

		public async Task OnLoadedAsync()
		{
			Log.Information("[ExcelMerger] View 已加载，开始初始化...");
			await Task.Yield();
		}

		private async Task ReadFromActiveExcelAsync()
		{
			await RunWithProgressAsync("正在连接活动的 Microsoft Excel...", async progress =>
			{
				var info = await _excelInteropService.DetectActiveExcelAsync(progress, CancellationToken.None);
				AddLog($"系统 - 成功连接到活动的 Microsoft Excel 版本 {info.Version}");
				Sheets.Clear();
				foreach (var sheet in info.SheetNames) Sheets.Add(sheet);
			});
		}

		private async Task PickFileAsync()
		{
			// 简化：直接从工作目录读取一个文件名，真实项目中使用 OpenFileDialog
			var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "sample.xlsx");
			AddLog($"[ExcelMerger] 用户操作 - 添加文件: {Path.GetFileName(path)}, 大小: {GetFileSizeKb(path)}");

			await RunWithProgressAsync("正在解析Excel...", async progress =>
			{
				var parsed = await _excelFileParserService.ParseWorkbookAsync(path, progress, CancellationToken.None);
				Sheets.Clear();
				foreach (var sheet in parsed.SheetNames) Sheets.Add(sheet);
				AddLog($"[ExcelMerger] 解析完成 - 文件: {Path.GetFileName(path)}, 工作表数: {parsed.SheetNames.Count}");
			});
		}

		private async Task StartProcessAsync()
		{
			await RunWithProgressAsync("正在执行匹配...", async progress =>
			{
				await _matchService.MatchAsync(progress, CancellationToken.None);
				AddLog("[ExcelMerger] 匹配与回写完成");
			});
		}

		private async Task RunWithProgressAsync(string startText, Func<IProgress<(int,string)>, Task> work)
		{
			ProgressText = startText;
			ProgressValue = 0;
			var progress = new Progress<(int percent, string message)>(t =>
			{
				ProgressValue = t.percent;
				ProgressText = t.message;
			});
			try
			{
				await Task.Run(() => work(progress));
			}
			catch (Exception ex)
			{
				AddLog($"错误: {ex.Message}");
			}
		}

		private void AddLog(string message)
		{
			Logs.Add(message);
			Log.Information(message);
		}

		private static string GetFileSizeKb(string path)
		{
			try { var fi = new FileInfo(path); return $"{fi.Length / 1024.0:F2} KB"; } catch { return "未知"; }
		}
	}
}