using System;
using System.Windows.Forms;

namespace YYToolsTest
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // 检查命令行参数，决定使用控制台模式还是界面模式
            bool useConsole = args.Length > 0 && args[0].ToLower() == "console";
            
            if (useConsole)
            {
                RunConsoleMode();
            }
            else
            {
                RunGuiMode();
            }
        }
        
        static void RunGuiMode()
        {
            try
            {
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动界面模式失败: " + ex.Message, "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        static void RunConsoleMode()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("YYTools 综合测试程序 v2.1 (控制台模式)");
            Console.WriteLine("========================================");
            
            try
            {
                // 测试1：创建COM对象
                Console.WriteLine("1. 测试创建COM对象...");
                object yyTools = Activator.CreateInstance(Type.GetTypeFromProgID("YYTools.ExcelAddin"));
                if (yyTools != null)
                {
                    Console.WriteLine("✓ COM对象创建成功");
                    
                    // 测试2：调用GetDetailedApplicationInfo方法
                    Console.WriteLine("\n2. 测试GetDetailedApplicationInfo方法...");
                    try
                    {
                        string info = (string)yyTools.GetType().InvokeMember("GetDetailedApplicationInfo",
                            System.Reflection.BindingFlags.InvokeMethod, null, yyTools, null);
                        Console.WriteLine("✓ GetDetailedApplicationInfo调用成功");
                        Console.WriteLine("详细信息:");
                        Console.WriteLine(info);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("✗ GetDetailedApplicationInfo调用失败: " + ex.Message);
                    }
                    
                    // 测试3：调用InstallMenu方法
                    Console.WriteLine("\n3. 测试InstallMenu方法...");
                    try
                    {
                        string result = (string)yyTools.GetType().InvokeMember("InstallMenu",
                            System.Reflection.BindingFlags.InvokeMethod, null, yyTools, null);
                        Console.WriteLine("✓ InstallMenu调用成功");
                        Console.WriteLine("结果: " + result);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("✗ InstallMenu调用失败: " + ex.Message);
                    }
                    
                    // 测试4：调用ShowMatchForm方法
                    Console.WriteLine("\n4. 测试ShowMatchForm方法...");
                    try
                    {
                        yyTools.GetType().InvokeMember("ShowMatchForm",
                            System.Reflection.BindingFlags.InvokeMethod, null, yyTools, null);
                        Console.WriteLine("✓ ShowMatchForm调用成功");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("✗ ShowMatchForm调用失败: " + ex.Message);
                    }
                    
                    // 测试5：调用RefreshMenu方法
                    Console.WriteLine("\n5. 测试RefreshMenu方法...");
                    try
                    {
                        yyTools.GetType().InvokeMember("RefreshMenu",
                            System.Reflection.BindingFlags.InvokeMethod, null, yyTools, null);
                        Console.WriteLine("✓ RefreshMenu调用成功");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("✗ RefreshMenu调用失败: " + ex.Message);
                    }
                    
                    // 测试6：获取简单应用程序信息
                    Console.WriteLine("\n6. 测试GetApplicationInfo方法...");
                    try
                    {
                        string info = (string)yyTools.GetType().InvokeMember("GetApplicationInfo",
                            System.Reflection.BindingFlags.InvokeMethod, null, yyTools, null);
                        Console.WriteLine("✓ GetApplicationInfo调用成功");
                        Console.WriteLine("基本信息: " + info);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("✗ GetApplicationInfo调用失败: " + ex.Message);
                    }
                    
                    // 清理
                    yyTools = null;
                }
                else
                {
                    Console.WriteLine("✗ COM对象创建失败");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("✗ 创建COM对象异常: " + ex.Message);
                Console.WriteLine("\n可能的原因:");
                Console.WriteLine("1. YYTools.dll未正确注册 - 运行 install_admin.bat");
                Console.WriteLine("2. 需要管理员权限 - 右键以管理员身份运行");
                Console.WriteLine("3. DLL文件路径不正确 - 检查 bin\\Debug\\YYTools.dll");
                Console.WriteLine("4. .NET Framework版本不兼容");
                
                Console.WriteLine("\n调试建议:");
                Console.WriteLine("1. 运行 check_registration.bat 检查注册状态");
                Console.WriteLine("2. 查看Windows事件查看器中的应用程序日志");
                Console.WriteLine("3. 确保WPS表格或Excel已启动并打开工作簿");
            }
            
            Console.WriteLine("\n========================================");
            Console.WriteLine("测试完成！");
            Console.WriteLine("\n如果所有测试通过，请:");
            Console.WriteLine("1. 打开WPS表格/Excel");
            Console.WriteLine("2. 查看工具栏是否有'YY工具'菜单");
            Console.WriteLine("3. 如果没有，在VBA中运行:");
            Console.WriteLine("   CreateObject(\"YYTools.ExcelAddin\").InstallMenu()");
            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }
    }
} 