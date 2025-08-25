using System;
using System.Windows.Forms;
using System.Diagnostics;
using YYTools;

class DebugTestProgram
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // 显示调试控制台
        AllocConsole();
        
        Console.WriteLine("=== YY运单匹配工具 - 调试测试程序 ===");
        Console.WriteLine("正在启动调试模式...");
        Console.WriteLine();

        try
        {
            Console.WriteLine("1. 检查进程状态...");
            CheckProcesses();
            Console.WriteLine();

            Console.WriteLine("2. 尝试连接Excel/WPS...");
            var app = ExcelAddin.GetExcelApplication();
            
            if (app == null)
            {
                Console.WriteLine("❌ 无法连接到Excel/WPS应用程序！");
                Console.WriteLine();
                Console.WriteLine("请确保：");
                Console.WriteLine("- WPS表格或Excel已启动");
                Console.WriteLine("- 至少打开一个工作簿文件");
                Console.WriteLine("- 文件没有处于保护模式");
                Console.WriteLine();
                Console.WriteLine("按任意键退出...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("✅ 成功连接到应用程序！");
            Console.WriteLine("应用程序名称: " + GetSafeProperty(app, "Name"));
            Console.WriteLine("应用程序版本: " + GetSafeProperty(app, "Version"));
            Console.WriteLine();

            Console.WriteLine("3. 检查工作簿...");
            var workbooks = ExcelAddin.GetOpenWorkbooks();
            Console.WriteLine("找到 " + workbooks.Count + " 个工作簿:");
            
            for (int i = 0; i < workbooks.Count; i++)
            {
                var wb = workbooks[i];
                Console.WriteLine("  [" + (i + 1) + "] " + wb.Name + " " + (wb.IsActive ? "★活动" : ""));
                
                // 检查工作表
                var sheets = ExcelAddin.GetWorksheetNames(wb.Workbook);
                Console.WriteLine("      包含 " + sheets.Count + " 个工作表: " + string.Join(", ", sheets.ToArray()));
            }
            Console.WriteLine();

            if (workbooks.Count == 0)
            {
                Console.WriteLine("❌ 没有找到打开的工作簿！");
                Console.WriteLine("请在WPS/Excel中打开包含数据的文件后再试。");
                Console.WriteLine();
                Console.WriteLine("按任意键退出...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("4. 启动匹配工具...");
            Console.WriteLine("如果工具正常启动，说明问题已解决！");
            Console.WriteLine();

            ExcelAddin.ShowMatchForm();
        }
        catch (Exception ex)
        {
            Console.WriteLine("❌ 发生异常：");
            Console.WriteLine(ex.ToString());
            Console.WriteLine();
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }

    static void CheckProcesses()
    {
        // 检查WPS进程
        var wpsProcesses = Process.GetProcessesByName("wps");
        Console.WriteLine("WPS进程数量: " + wpsProcesses.Length);
        foreach (var proc in wpsProcesses)
        {
            Console.WriteLine("  - WPS进程: " + proc.ProcessName + " (PID: " + proc.Id + ")");
        }

        // 检查Excel进程
        var excelProcesses = Process.GetProcessesByName("excel");
        Console.WriteLine("Excel进程数量: " + excelProcesses.Length);
        foreach (var proc in excelProcesses)
        {
            Console.WriteLine("  - Excel进程: " + proc.ProcessName + " (PID: " + proc.Id + ")");
        }
    }

    static string GetSafeProperty(object obj, string propertyName)
    {
        try
        {
            if (obj == null) return "null";
            var prop = obj.GetType().GetProperty(propertyName);
            if (prop == null) return "属性不存在";
            var value = prop.GetValue(obj, null);
            return value != null ? value.ToString() : "null";
        }
        catch (Exception ex)
        {
            return "错误: " + ex.Message;
        }
    }

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    static extern bool AllocConsole();
} 