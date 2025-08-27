// --- 文件 3: ExcelAddin.cs ---
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    [ComVisible(true)]
    [Guid("12345678-1234-5678-9ABC-123456789ABC")]
    [ProgId("YYTools.ExcelAddin")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ExcelAddin
    {
        private static Excel.Application application;
        public static Excel.Application Application => application;

        [ComVisible(true)]
        public static void ShowMatchForm()
        {
            try
            {
                var matchForm = new MatchForm();
                matchForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动运单匹配工具失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static Excel.Application GetExcelApplication()
        {
            try
            {
                if (IsApplicationValid(application)) return application;

                Excel.Application app = null;
                string[] wpsProgIds = { "Ket.Application", "WPS.Application", "Kingsoft.Application", "ET.Application" };
                string[] excelProgIds = { "Excel.Application" };

                foreach (string progId in wpsProgIds)
                {
                    app = TryGetActiveApp(progId);
                    if (IsApplicationValid(app)) { application = app; return app; }
                }

                foreach (string progId in excelProgIds)
                {
                    app = TryGetActiveApp(progId);
                    if (IsApplicationValid(app)) { application = app; return app; }
                }

                System.Threading.Thread.Sleep(300);
                foreach (string progId in wpsProgIds)
                {
                    app = TryGetActiveApp(progId);
                    if (IsApplicationValid(app)) { application = app; return app; }
                }

                app = TryGetFromROTByKeywords(new[] { "Ket.Application", "ET.Application", "WPS.Application", "Kingsoft.Application", "Excel.Application" });
                if (IsApplicationValid(app)) { application = app; return app; }

                try
                {
                    Type excelType = Type.GetTypeFromProgID("Excel.Application");
                    app = Activator.CreateInstance(excelType) as Excel.Application;
                    if (IsApplicationValid(app)) { application = app; return app; }
                }
                catch { }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Application TryGetActiveApp(string progId)
        {
            try { return (Excel.Application)Marshal.GetActiveObject(progId); }
            catch { return null; }
        }

        private static bool IsApplicationValid(Excel.Application app)
        {
            try
            {
                if (app == null) return false;
                _ = app.Name;
                return true;
            }
            catch { return false; }
        }

        public static bool HasOpenWorkbooks(Excel.Application app)
        {
            try
            {
                if (!IsApplicationValid(app)) return false;
                return app.Workbooks.Count > 0;
            }
            catch { return false; }
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private static Excel.Application TryGetFromROTByKeywords(string[] keywords)
        {
            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMoniker = null;
            try
            {
                if (CreateBindCtx(0, out bindCtx) != 0) return null;
                bindCtx.GetRunningObjectTable(out rot);
                if (rot == null) return null;
                rot.EnumRunning(out enumMoniker);
                if (enumMoniker == null) return null;
                enumMoniker.Reset();
                IMoniker[] monikers = new IMoniker[1];
                while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    string displayName;
                    monikers[0].GetDisplayName(bindCtx, null, out displayName);
                    if (keywords.Any(keyword => displayName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        object obj;
                        rot.GetObject(monikers[0], out obj);
                        var app = obj as Excel.Application;
                        if (IsApplicationValid(app)) return app;
                    }
                }
            }
            catch { }
            finally
            {
                if (enumMoniker != null) Marshal.ReleaseComObject(enumMoniker);
                if (rot != null) Marshal.ReleaseComObject(rot);
                if (bindCtx != null) Marshal.ReleaseComObject(bindCtx);
            }
            return null;
        }

        public static List<WorkbookInfo> GetOpenWorkbooks()
        {
            var result = new List<WorkbookInfo>();
            try
            {
                Excel.Application app = GetExcelApplication();
                if (!IsApplicationValid(app)) return result;

                Excel.Workbook activeWorkbook = null;
                try { activeWorkbook = app.ActiveWorkbook; } catch { }

                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    if (wb != null && !string.IsNullOrEmpty(wb.Name))
                    {
                        bool isActive = activeWorkbook != null && string.Equals(wb.Name, activeWorkbook.Name, StringComparison.OrdinalIgnoreCase);
                        result.Add(new WorkbookInfo { Name = wb.Name, Workbook = wb, IsActive = isActive });
                    }
                }
                if (result.Count > 0 && result.All(r => !r.IsActive))
                {
                    result[0].IsActive = true;
                }
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"获取打开的工作簿列表失败: {ex.Message}", LogLevel.Error);
            }
            return result;
        }

        public static List<string> GetWorksheetNames(Excel.Workbook workbook)
        {
            var list = new List<string>();
            try
            {
                if (workbook?.Worksheets != null)
                {
                    foreach (Excel.Worksheet ws in workbook.Worksheets)
                    {
                        if (ws != null) list.Add(ws.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"获取工作表名称失败: {ex.Message}", LogLevel.Error);
            }
            return list;
        }

        public static Excel.Workbook LoadWorkbookFromFile(string filePath)
        {
            try
            {
                var app = GetExcelApplication();
                if (app == null) throw new InvalidOperationException("无法连接到WPS或Excel应用程序。");

                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    if (string.Equals(wb.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        wb.Activate();
                        return wb;
                    }
                }

                Excel.Workbook workbook = app.Workbooks.Open(filePath);
                app.Visible = true;
                workbook.Activate();
                return workbook;
            }
            catch (Exception ex)
            {
                MatchService.WriteLog($"从文件加载工作簿失败: {filePath}. 错误: {ex.Message}", LogLevel.Error);
                return null;
            }
        }

        [ComRegisterFunction]
        public static void RegisterFunction(Type type) { }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type) { }
    }
}