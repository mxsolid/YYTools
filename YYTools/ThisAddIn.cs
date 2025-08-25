using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace YYTools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 插件启动时的初始化代码
            try
            {
                // 记录插件启动
                System.Diagnostics.Debug.WriteLine("YY运单匹配工具已启动");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"插件启动失败：{ex.Message}", "错误", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 插件关闭时的清理代码
            try
            {
                System.Diagnostics.Debug.WriteLine("YY运单匹配工具已关闭");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"插件关闭时出错：{ex.Message}");
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
} 