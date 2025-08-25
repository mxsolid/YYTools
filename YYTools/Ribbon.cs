using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace YYTools
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            // 功能区加载完成
        }

        /// <summary>
        /// 运单匹配按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMatch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 检查是否有打开的工作簿
                if (Globals.ThisAddIn.Application.Workbooks.Count == 0)
                {
                    MessageBox.Show("请先打开一个Excel文件！", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 创建并显示匹配窗体
                MatchForm form = new MatchForm();
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动运单匹配工具时发生错误：\n{ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 关于按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string aboutText = "YY运单匹配工具 v1.0\n\n" +
                                   "功能：快速匹配发货明细和账单明细中的运单号\n" +
                                   "开发：YYTools\n" +
                                   "版权：© 2025 YYTools 保留所有权利\n\n" +
                                   "使用方法：\n" +
                                   "1. 打开包含发货明细和账单明细的Excel文件\n" +
                                   "2. 点击'运单匹配'按钮\n" +
                                   "3. 选择对应的工作表和列\n" +
                                   "4. 点击'开始匹配'完成数据填充";

                MessageBox.Show(aboutText, "关于 YY运单匹配工具", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"显示关于信息时发生错误：\n{ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
} 