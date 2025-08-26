using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// Excel插件类 - 简化版WPS优先检测
    /// </summary>
    [ComVisible(true)]
    public class ExcelAddin
    {
        // ... (Your other methods like ShowMatchForm, ShowSettings remain the same) ...

        /// <summary>
        /// 显示匹配窗体
        /// </summary>
        public void ShowMatchForm()
        {
            try
            {
                var form = new MatchForm();
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("显示匹配窗体失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 显示设置窗体
        /// </summary>
        public void ShowSettings()
        {
            try
            {
                var settingsForm = new SettingsForm();
                settingsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("显示设置窗体失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ====================================================================
        // ===            IMPROVED METHOD STARTS HERE                       ===
        // ====================================================================

        /// <summary>
        /// 获取Excel应用程序实例 - (新) 更稳定可靠的WPS/Excel检测方法
        /// </summary>
        public static Excel.Application GetExcelApplication()
        {
            // NEW: Define ProgIDs for WPS and Excel. We prioritize WPS by putting it first.
            var progIds = new[] { "Ket.Application","ET.Application", "Excel.Application" };

            foreach (var progId in progIds)
            {
                try
                {
                    // Try to get a running instance using its ProgID
                    var app = (Excel.Application)Marshal.GetActiveObject(progId);

                    // MORE RELIABLE CHECK:
                    // Ensure the instance is valid, visible to the user, and has open workbooks.
                    // This prevents connecting to hidden or background Excel processes.
                    if (app != null && app.Visible && app.Workbooks.Count > 0)
                    {
                        Debug.WriteLine($"Successfully connected to {progId}");
                        return app; // Found a valid, running instance. Return it immediately.
                    }
                }
                catch (COMException)
                {
                    // This is expected if an application with the given ProgID is not running.
                    // We can safely ignore it and try the next one in the list.
                    Debug.WriteLine($"{progId} not found in Running Object Table. Trying next...");
                }
                catch (Exception ex)
                {
                    // Catch any other unexpected errors during connection.
                    MessageBox.Show($"连接到 {progId} 时发生意外错误: {ex.Message}", "连接错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // If the loop completes without finding any suitable instance
            Debug.WriteLine("No suitable running instance of WPS or Excel found.");
            return null;
        }

        // ====================================================================
        // ===            IMPROVED METHOD ENDS HERE                         ===
        // ====================================================================


        /// <summary>
        /// 检查是否有打开的工作簿 - 简单检测
        /// </summary>
        private static bool HasOpenWorkbooks(Excel.Application app)
        {
            try
            {
                if (app == null) return false;

                var workbooks = app.Workbooks;
                if (workbooks == null) return false;

                return workbooks.Count > 0;
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// 获取所有工作簿
        /// </summary>
        public static List<Excel.Workbook> GetWorkbooks()
        {
            try
            {
                var app = GetExcelApplication();
                if (app == null) return new List<Excel.Workbook>();

                var workbooks = new List<Excel.Workbook>();
                foreach (Excel.Workbook workbook in app.Workbooks)
                {
                    workbooks.Add(workbook);
                }
                return workbooks;
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取工作簿列表失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new List<Excel.Workbook>();
            }
        }

        /// <summary>
        /// 获取工作表名称列表
        /// </summary>
        public static List<string> GetWorksheetNames(Excel.Workbook workbook)
        {
            try
            {
                var names = new List<string>();
                if (workbook == null) return names;

                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    names.Add(worksheet.Name);
                }
                return names;
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取工作表列表失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new List<string>();
            }
        }

        /// <summary>
        /// 获取工作表
        /// </summary>
        public static Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string worksheetName)
        {
            try
            {
                if (workbook == null || string.IsNullOrEmpty(worksheetName)) return null;

                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name == worksheetName)
                    {
                        return worksheet;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取工作表失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// 获取列标题
        /// </summary>
        public static List<string> GetColumnHeaders(Excel.Worksheet worksheet)
        {
            try
            {
                var headers = new List<string>();
                if (worksheet == null) return headers;

                var usedRange = worksheet.UsedRange;
                if (usedRange == null) return headers;

                int columnCount = usedRange.Columns.Count;
                for (int i = 1; i <= columnCount; i++)
                {
                    var cell = (Excel.Range)worksheet.Cells[1, i];
                    var value = cell.Value2;
                    headers.Add(value != null ? value.ToString() : "列" + i);
                }

                return headers;
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取列标题失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new List<string>();
            }
        }
    }
}