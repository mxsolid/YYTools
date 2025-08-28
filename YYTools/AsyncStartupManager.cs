using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// 启动阶段的异步预热管理器：预解析已打开的工作簿/工作表，填充缓存，避免首次交互卡顿
    /// </summary>
    public static class AsyncStartupManager
    {
        /// <summary>
        /// 预热Excel数据到缓存（不阻塞UI）
        /// </summary>
        public static Task WarmUpAsync(CancellationToken token, IProgress<TaskProgress> progress)
        {
            return Task.Run(async () =>
            {
                DateTime start = DateTime.Now;
                try
                {
                    progress?.Report(new TaskProgress(5, "正在检测已打开的Excel/WPS工作簿..."));
                    var app = ExcelAddin.GetExcelApplication();
                    if (app == null || !ExcelAddin.HasOpenWorkbooks(app))
                    {
                        progress?.Report(new TaskProgress(100, "未检测到打开的Excel/WPS文件，跳过预热"));
                        return;
                    }

                    List<WorkbookInfo> workbooks = ExcelAddin.GetOpenWorkbooks();
                    int total = 0;
                    foreach (var wb in workbooks)
                    {
                        if (token.IsCancellationRequested) token.ThrowIfCancellationRequested();

                        progress?.Report(new TaskProgress(Math.Min(15 + total, 90), $"读取工作表列表: {wb.Name}"));
                        var sheetNames = DataManager.GetSheetNames(wb.Workbook);

                        int count = 0;
                        foreach (var sheetName in sheetNames)
                        {
                            if (token.IsCancellationRequested) token.ThrowIfCancellationRequested();

                            try
                            {
                                var ws = wb.Workbook.Worksheets[sheetName] as Excel.Worksheet;
                                if (ws == null) continue;
                                // 解析列信息（限定前50行，由 SmartColumnService 控制）
                                DataManager.GetColumnInfos(ws);
                                count++;
                                total++;
                                int pct = Math.Min(90, 20 + total % 70);
                                progress?.Report(new TaskProgress(pct, $"预热列信息: {wb.Name} / {sheetName}"));
                            }
                            catch (Exception exSheet)
                            {
                                Logger.LogWarning($"预热工作表失败: {wb.Name} / {sheetName} - {exSheet.Message}");
                            }
                            await Task.Yield();
                        }

                        Logger.LogExcelOperation("预热缓存", wb.Name, "", 0, 0);
                    }

                    var elapsed = DateTime.Now - start;
                    Logger.LogPerformance("启动预热完成", elapsed, $"处理工作簿: {workbooks.Count}");
                    progress?.Report(new TaskProgress(100, "预热完成"));
                }
                catch (OperationCanceledException)
                {
                    progress?.Report(new TaskProgress(100, "预热已取消"));
                }
                catch (Exception ex)
                {
                    Logger.LogError("启动预热失败", ex);
                    progress?.Report(new TaskProgress(100, $"预热失败: {ex.Message}"));
                }
            }, token);
        }
    }
}