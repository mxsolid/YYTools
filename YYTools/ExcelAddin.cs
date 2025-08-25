using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace YYTools
{
    /// <summary>
    /// Excel/WPS COM加载项 - 增强版本，支持独立菜单和直接调用
    /// </summary>
    [ComVisible(true)]
    [Guid("12345678-1234-5678-9ABC-123456789ABC")]
    [ProgId("YYTools.ExcelAddin")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ExcelAddin
    {
        private static Excel.Application application;

        public static Excel.Application Application
        {
            get { return application; }
        }

        /// <summary>
        /// 显示匹配窗体 - 直接调用MatchForm
        /// </summary>
        [ComVisible(true)]
        public static void ShowMatchForm()
        {
            try
            {
                // 获取Excel应用程序实例
                application = GetExcelApplication();
                if (application == null)
                {
                    MessageBox.Show("无法连接到WPS表格或Excel。\n\n请确认：\n1. 已启动WPS表格或Excel\n2. 至少打开一个工作簿文件\n3. 文件未处于受保护或只读限制模式", 
                        "连接失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查工作簿
                if (!HasOpenWorkbooks(application))
                {
                    MessageBox.Show("已连接到WPS/Excel，但未检测到打开的工作簿。\n\n请打开包含数据的表格后重试。", 
                        "未发现工作簿", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 直接创建并显示匹配窗体
                var matchForm = new MatchForm();
                matchForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动运单匹配工具失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 显示设置窗体 - 直接调用SettingsForm
        /// </summary>
        [ComVisible(true)]
        public static void ShowSettings()
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
        /// 获取Excel/WPS应用程序实例 - WPS优先版本
        /// </summary>
        public static Excel.Application GetExcelApplication()
        {
            try
            {
                Excel.Application app = null;
                
                // WPS优先策略 - 按优先级排序
                string[] wpsProgIds = { "Ket.Application", "WPS.Application", "Kingsoft.Application", "ET.Application" };
                string[] excelProgIds = { "Excel.Application" };
                
                // 1. 尝试WPS各种ProgID
                foreach (string progId in wpsProgIds)
                {
                    app = TryGetActiveApp(progId);
                    if (IsApplicationValid(app)) return app;
                }
                
                // 2. 尝试Excel
                foreach (string progId in excelProgIds)
                {
                    app = TryGetActiveApp(progId);
                    if (IsApplicationValid(app)) return app;
                }

                // 3. 短暂等待后再试WPS（处理启动延迟）
                System.Threading.Thread.Sleep(300);
                foreach (string progId in wpsProgIds)
                {
                    app = TryGetActiveApp(progId);
                    if (IsApplicationValid(app)) return app;
                }

                // 4. ROT兜底策略
                app = TryGetFromROTByKeywords(new string[] { 
                    "Ket.Application", "ET.Application", "WPS.Application", 
                    "Kingsoft.Application", "Excel.Application" 
                });
                if (IsApplicationValid(app)) return app;

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Application TryGetActiveApp(string progId)
        {
            try
            {
                return (Excel.Application)Marshal.GetActiveObject(progId);
            }
            catch
            {
                return null;
            }
        }

        private static bool IsApplicationValid(Excel.Application app)
        {
            try
            {
                if (app == null) return false;
                string name = app.Name;
                return !string.IsNullOrEmpty(name) && app.Workbooks != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 重写的工作簿检测方法 - 更可靠的检测逻辑
        /// </summary>
        private static bool HasOpenWorkbooks(Excel.Application app)
        {
            try
            {
                if (app == null) return false;
                
                // 方法1：直接获取数量
                try
                {
                    if (app.Workbooks.Count > 0) return true;
                }
                catch { }
                
                // 方法2：尝试访问ActiveWorkbook
                try
                {
                    if (app.ActiveWorkbook != null && !string.IsNullOrEmpty(app.ActiveWorkbook.Name)) 
                        return true;
                }
                catch { }
                
                // 方法3：遍历检测
                try
                {
                    foreach (Excel.Workbook wb in app.Workbooks)
                    {
                        if (wb != null && !string.IsNullOrEmpty(wb.Name)) return true;
                    }
                }
                catch { }
                
                return false;
            }
            catch
            {
                return false;
            }
        }

        // ========== ROT 兜底 ========== //
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private static Excel.Application TryGetFromROTByKeywords(string[] keywords)
        {
            try
            {
                IBindCtx bindCtx;
                if (CreateBindCtx(0, out bindCtx) != 0 || bindCtx == null) return null;

                IRunningObjectTable rot;
                bindCtx.GetRunningObjectTable(out rot);
                if (rot == null) return null;

                IEnumMoniker enumMoniker;
                rot.EnumRunning(out enumMoniker);
                enumMoniker.Reset();

                IMoniker[] monikers = new IMoniker[1];
                IntPtr fetched = IntPtr.Zero;

                while (enumMoniker.Next(1, monikers, fetched) == 0)
                {
                    string displayName = null;
                    try
                    {
                        monikers[0].GetDisplayName(bindCtx, null, out displayName);
                    }
                    catch
                    {
                        displayName = null;
                    }

                    if (!string.IsNullOrEmpty(displayName))
                    {
                        for (int i = 0; i < keywords.Length; i++)
                        {
                            if (displayName.IndexOf(keywords[i], StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                try
                                {
                                    object obj;
                                    rot.GetObject(monikers[0], out obj);
                                    var app = obj as Excel.Application;
                                    if (IsApplicationValid(app)) return app;
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 获取打开的工作簿列表 - 重写版本，默认选中当前激活文件
        /// </summary>
        public static List<WorkbookInfo> GetOpenWorkbooks()
        {
            List<WorkbookInfo> result = new List<WorkbookInfo>();
            try
            {
                Excel.Application app = GetExcelApplication();
                if (!IsApplicationValid(app)) return result;

                Excel.Workbook activeWorkbook = null;
                try { activeWorkbook = app.ActiveWorkbook; } catch { activeWorkbook = null; }

                // 遍历所有工作簿
                try
                {
                    foreach (Excel.Workbook wb in app.Workbooks)
                    {
                        if (wb != null && !string.IsNullOrEmpty(wb.Name))
                        {
                            bool isActive = (activeWorkbook != null && 
                                string.Equals(wb.Name, activeWorkbook.Name, StringComparison.OrdinalIgnoreCase));
                            result.Add(new WorkbookInfo { Name = wb.Name, Workbook = wb, IsActive = isActive });
                        }
                    }
                }
                catch { }

                // 如果没有发现激活的工作簿，但有工作簿列表，则将第一个设为激活
                if (result.Count > 0)
                {
                    bool hasActive = false;
                    for (int i = 0; i < result.Count; i++)
                    {
                        if (result[i].IsActive)
                        {
                            hasActive = true;
                            break;
                        }
                    }
                    if (!hasActive)
                    {
                        result[0].IsActive = true;
                    }
                }
            }
            catch { }
            return result;
        }

        /// <summary>
        /// 获取工作表名称列表
        /// </summary>
        public static List<string> GetWorksheetNames(Excel.Workbook workbook)
        {
            List<string> list = new List<string>();
            try
            {
                if (workbook?.Worksheets != null)
                {
                    foreach (Excel.Worksheet ws in workbook.Worksheets)
                    {
                        if (ws != null && !string.IsNullOrEmpty(ws.Name))
                        {
                            list.Add(ws.Name);
                        }
                    }
                }
            }
            catch { }
            return list;
        }

        /// <summary>
        /// 创建YY工具独立菜单 - 类似方方格子的实现
        /// </summary>
        [ComVisible(true)]
        public static void CreateYYToolsMenu()
        {
            try
            {
                Excel.Application app = GetExcelApplication();
                if (!IsApplicationValid(app)) return;

                // 删除已存在的菜单
                try
                {
                    var existingMenu = app.CommandBars["YY工具"];
                    if (existingMenu != null)
                    {
                        existingMenu.Delete();
                    }
                }
                catch { }

                // 创建新的独立菜单栏
                Microsoft.Office.Core.CommandBar menuBar = null;
                try
                {
                    // 创建为菜单栏类型 (msoBarTypeMenuBar = 1)
                    menuBar = app.CommandBars.Add("YY工具", Microsoft.Office.Core.MsoBarPosition.msoBarTop, false, true);
                    menuBar.Visible = true;
                    menuBar.Position = Microsoft.Office.Core.MsoBarPosition.msoBarTop;
                    
                    // 添加运单匹配工具按钮
                    var matchButton = (Microsoft.Office.Core.CommandBarButton)menuBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
                    matchButton.Caption = "运单匹配工具";
                    matchButton.TooltipText = "打开运单匹配配置窗体";
                    matchButton.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                    matchButton.FaceId = 162; // 使用内置图标
                    
                    // 重要：直接绑定C#方法，而不是VBA宏
                    matchButton.OnAction = "YYToolsShowMatchForm";
                    
                    // 添加工具设置按钮
                    var settingsButton = (Microsoft.Office.Core.CommandBarButton)menuBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
                    settingsButton.Caption = "工具设置";
                    settingsButton.TooltipText = "打开工具设置窗体";
                    settingsButton.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                    settingsButton.FaceId = 642; // 使用内置图标
                    
                    // 绑定设置方法
                    settingsButton.OnAction = "YYToolsShowSettings";

                    // 添加分隔符
                    menuBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    
                    // 添加关于按钮
                    var aboutButton = (Microsoft.Office.Core.CommandBarButton)menuBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
                    aboutButton.Caption = "关于YY工具";
                    aboutButton.TooltipText = "关于YY运单匹配工具";
                    aboutButton.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                    aboutButton.FaceId = 487; // 使用内置图标
                    aboutButton.OnAction = "YYToolsShowAbout";

                }
                catch (Exception ex)
                {
                    // 如果菜单栏创建失败，尝试添加到现有菜单栏
                    try
                    {
                        var workbookMenu = app.CommandBars["Worksheet Menu Bar"];
                        if (workbookMenu != null)
                        {
                            // 添加到现有菜单栏
                            var yyToolsMenu = (Microsoft.Office.Core.CommandBarPopup)workbookMenu.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            yyToolsMenu.Caption = "YY工具";
                            
                            // 添加子菜单项
                            var matchMenuItem = (Microsoft.Office.Core.CommandBarButton)yyToolsMenu.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            matchMenuItem.Caption = "运单匹配工具";
                            matchMenuItem.OnAction = "YYToolsShowMatchForm";
                            matchMenuItem.FaceId = 162;

                            var settingsMenuItem = (Microsoft.Office.Core.CommandBarButton)yyToolsMenu.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            settingsMenuItem.Caption = "工具设置";
                            settingsMenuItem.OnAction = "YYToolsShowSettings";
                            settingsMenuItem.FaceId = 642;
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// 显示关于信息
        /// </summary>
        [ComVisible(true)]
        public static void ShowAbout()
        {
            try
            {
                string aboutInfo = "YY运单匹配工具 v2.1\n\n" +
                                 "功能特点：\n" +
                                 "• 智能运单匹配算法\n" +
                                 "• 支持多工作簿操作\n" +
                                 "• 自动列选择和验证\n" +
                                 "• 批量数据处理\n\n" +
                                 "适用于：WPS表格、Microsoft Excel\n\n" +
                                 "开发：YY工具团队";
                
                MessageBox.Show(aboutInfo, "关于YY工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("显示关于信息失败：" + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 强制刷新菜单 - 公开方法供外部调用
        /// </summary>
        [ComVisible(true)]
        public static void RefreshMenu()
        {
            try
            {
                CreateYYToolsMenu();
            }
            catch { }
        }

        /// <summary>
        /// 手动安装菜单 - 公开方法，可从VBA或测试程序调用
        /// </summary>
        [ComVisible(true)]
        public static string InstallMenu()
        {
            try
            {
                Excel.Application app = GetExcelApplication();
                if (!IsApplicationValid(app))
                {
                    return "错误：无法连接到WPS表格或Excel应用程序";
                }

                CreateYYToolsMenu();
                return "YY工具菜单安装成功！请查看WPS/Excel工具栏";
            }
            catch (Exception ex)
            {
                return "菜单安装失败：" + ex.Message;
            }
        }

        /// <summary>
        /// 检查应用程序兼容性
        /// </summary>
        [ComVisible(true)]
        public static string GetApplicationInfo()
        {
            try
            {
                Excel.Application app = GetExcelApplication();
                if (!IsApplicationValid(app)) return "未检测到WPS或Excel应用程序";
                
                string appName = app.Name;
                int workbookCount = app.Workbooks?.Count ?? 0;
                
                return string.Format("应用程序: {0}, 工作簿数量: {1}", appName, workbookCount);
            }
            catch (Exception ex)
            {
                return "获取应用程序信息失败: " + ex.Message;
            }
        }

        /// <summary>
        /// 获取详细的应用程序信息 - 增强版本
        /// </summary>
        [ComVisible(true)]
        public static string GetDetailedApplicationInfo()
        {
            try
            {
                Excel.Application app = GetExcelApplication();
                if (!IsApplicationValid(app)) 
                {
                    return "状态：未检测到WPS或Excel应用程序\n" +
                           "建议：\n" +
                           "1. 启动WPS表格或Excel\n" +
                           "2. 打开至少一个工作簿文件\n" +
                           "3. 确保文件未处于保护模式";
                }
                
                string appName = app.Name;
                int workbookCount = app.Workbooks?.Count ?? 0;
                string workbookNames = "";
                
                try
                {
                    // 获取工作簿名称
                    int count = 0;
                    foreach (Excel.Workbook wb in app.Workbooks)
                    {
                        if (wb != null && !string.IsNullOrEmpty(wb.Name))
                        {
                            if (count > 0) workbookNames += ", ";
                            workbookNames += wb.Name;
                            count++;
                            if (count >= 5) // 最多显示5个
                            {
                                workbookNames += "...";
                                break;
                            }
                        }
                    }
                }
                catch { }
                
                string result = "=== YYTools 应用程序信息 ===\n";
                result += "应用程序: " + appName + "\n";
                result += "工作簿数量: " + workbookCount + "\n";
                if (!string.IsNullOrEmpty(workbookNames))
                {
                    result += "工作簿列表: " + workbookNames + "\n";
                }
                result += "连接状态: 正常\n";
                result += "COM对象: 已创建\n";
                result += "建议操作: 可以调用 InstallMenu() 安装菜单";
                
                return result;
            }
            catch (Exception ex)
            {
                return "获取应用程序信息失败: " + ex.Message;
            }
        }

        /// <summary>
        /// COM注册
        /// </summary>
        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            try 
            { 
                System.Diagnostics.Debug.WriteLine("YYTools COM组件已注册");
            } 
            catch { }
        }

        /// <summary>
        /// COM反注册
        /// </summary>
        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                // 清理菜单
                Excel.Application app = GetExcelApplication();
                if (IsApplicationValid(app))
                {
                    try
                    {
                        var existing = app.CommandBars["YY工具"];
                        if (existing != null)
                        {
                            existing.Delete();
                        }
                    }
                    catch { }
                }
                System.Diagnostics.Debug.WriteLine("YYTools COM组件已反注册");
            }
            catch { }
        }
    }

    /// <summary>
    /// 工作簿信息类
    /// </summary>
    public class WorkbookInfo
    {
        public string Name { get; set; }
        public Excel.Workbook Workbook { get; set; }
        public bool IsActive { get; set; }
    }

    // ========== 全局VBA可调用方法 ========== //

    /// <summary>
    /// VBA可调用的全局方法 - 显示匹配窗体
    /// </summary>
    [ComVisible(true)]
    public class YYToolsGlobalMethods
    {
        [ComVisible(true)]
        public static void YYToolsShowMatchForm()
        {
            ExcelAddin.ShowMatchForm();
        }

        [ComVisible(true)]
        public static void YYToolsShowSettings()
        {
            ExcelAddin.ShowSettings();
        }

        [ComVisible(true)]
        public static void YYToolsShowAbout()
        {
            ExcelAddin.ShowAbout();
        }
    }
} 