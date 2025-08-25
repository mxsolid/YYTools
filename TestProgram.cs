using System;
using System.Windows.Forms;

namespace YYToolsTest
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                Console.WriteLine("启动 YY运单匹配工具...");
                Console.WriteLine("版本: v1.5 - 多文件支持，高性能优化");
                Console.WriteLine("========================================");
                
                // 调用运单匹配工具
                YYTools.ExcelAddin.ShowMatchForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("程序启动失败：\n{0}\n\n请确保已安装.NET Framework 4.8和Office组件", ex.Message), 
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine("错误: " + ex.Message);
            }
            
            Console.WriteLine("程序结束");
        }
    }
} 