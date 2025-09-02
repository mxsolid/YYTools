using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools.Services
{
	/// <summary>
	/// 通过 COM 与活动的 Microsoft Excel 交互。注意释放 COM 对象以避免进程残留。
	/// </summary>
	public class ExcelInteropService
	{
		public sealed class ActiveExcelInfo
		{
			public string Version { get; init; } = "";
			public List<string> SheetNames { get; init; } = new();
		}

		public async Task<ActiveExcelInfo> DetectActiveExcelAsync(IProgress<(int,string)>? progress, CancellationToken ct)
		{
			return await Task.Run(() =>
			{
				progress?.Report((5, "尝试连接 Excel 实例..."));
				Excel.Application? app = null;
				Excel.Workbook? wb = null;
				try
				{
					app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
					progress?.Report((20, "已获取 Excel 应用程序"));
					wb = app.ActiveWorkbook;
					var info = new ActiveExcelInfo { Version = app.Version };
					if (wb != null)
					{
						foreach (Excel.Worksheet ws in wb.Worksheets)
						{
							info.SheetNames.Add(ws.Name);
						}
					}
					progress?.Report((80, "已获取工作表信息"));
					return info;
				}
				finally
				{
					if (wb != null) Marshal.ReleaseComObject(wb);
					if (app != null) Marshal.ReleaseComObject(app);
					GC.Collect();
					GC.WaitForPendingFinalizers();
					progress?.Report((100, "完成"));
				}
			}, ct);
		}
	}
}