using System;
using System.Collections.Generic;
using System.Linq;
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

        /// <summary>
        /// 获取Excel应用程序实例 - WPS优先，简单有效
        /// </summary>
        public static Excel.Application GetExcelApplication()
        {
            try
            {
                // 1. 优先尝试WPS表格 (使用正确的ProgID)
                try
                {
                    var wpsApp = (Excel.Application)Marshal.GetActiveObject("Ket.Application");
                    if (wpsApp != null && HasOpenWorkbooks(wpsApp))
                    {
                        return wpsApp;
                    }
                }
                catch { }

                // 2. 尝试传统Excel
                try
                {
                    var excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    if (excelApp != null && HasOpenWorkbooks(excelApp))
                    {
                        return excelApp;
                    }
                }
                catch { }

                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("连接应用程序时出错: " + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

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