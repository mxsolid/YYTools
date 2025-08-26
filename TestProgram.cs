using System;
using System.Windows.Forms;

namespace YYToolsUltimate
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
                
                // 显示启动消息
                DialogResult startResult = MessageBox.Show(
                    "YY运单匹配工具 v1.5 - 终极性能优化版\n\n" +
                    "🚀 核心特性:\n" +
                    "• 50秒→5-8秒性能革命 (终极批量写入算法)\n" +
                    "• 多文件实时切换支持\n" +
                    "• 高分辨率屏幕完美适配\n" +
                    "• 智能任务管理 (停止/继续)\n" +
                    "• 详细任务总结报告\n\n" +
                    "准备启动工具吗？\n\n" +
                    "确保: 已在WPS表格中打开数据文件",
                    "YY运单匹配工具 - 启动确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (startResult == DialogResult.Yes)
                {
                    // 启动主窗体
                    var matchForm = new YYTools.MatchForm();
                    Application.Run(matchForm);
                }
                else
                {
                    MessageBox.Show("已取消启动，请在准备好数据文件后重新运行。", 
                        "启动取消", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = string.Format(
                    "程序启动失败\n\n" +
                    "错误详情: {0}\n\n" +
                    "可能原因:\n" +
                    "• 缺少.NET Framework 4.0运行时\n" +
                    "• YYTools.dll文件缺失或损坏\n" +
                    "• Office组件未正确安装\n\n" +
                    "建议:\n" +
                    "• 确保YYTools.dll在同一目录\n" +
                    "• 重新下载完整安装包\n" +
                    "• 以管理员身份运行",
                    ex.Message);
                
                MessageBox.Show(errorMsg, "启动错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
