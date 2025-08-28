using System;
using System.Drawing;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 用户引导窗体
    /// </summary>
    public partial class UserGuideForm : Form
    {
        private TabControl tabControl;
        private TabPage quickStartTab;
        private TabPage detailedGuideTab;
        private TabPage faqTab;
        private TabPage aboutTab;

        public UserGuideForm()
        {
            InitializeComponent();
            InitializeGuideContent();
            ApplyUIEnhancement();
        }

        private void InitializeComponent()
        {
            this.Text = "YY工具使用指南";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
        }

        private void InitializeGuideContent()
        {
            // 创建主控件
            tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;

            // 快速开始标签页
            CreateQuickStartTab();

            // 详细指南标签页
            CreateDetailedGuideTab();

            // 常见问题标签页
            CreateFaqTab();

            // 关于标签页
            CreateAboutTab();

            // 添加标签页到控件
            tabControl.TabPages.Add(quickStartTab);
            tabControl.TabPages.Add(detailedGuideTab);
            tabControl.TabPages.Add(faqTab);
            tabControl.TabPages.Add(aboutTab);

            // 添加关闭按钮
            var closeButton = new Button
            {
                Text = "关闭",
                Size = new Size(80, 30),
                Location = new Point(this.Width - 100, this.Height - 50),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            closeButton.Click += (s, e) => this.Close();

            // 添加控件到窗体
            this.Controls.Add(tabControl);
            this.Controls.Add(closeButton);
        }

        private void CreateQuickStartTab()
        {
            quickStartTab = new TabPage("快速开始");
            quickStartTab.BackColor = Color.White;

            var richTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("微软雅黑", 10F)
            };

            richTextBox.Text = @"快速开始指南

第一步：准备工作
• 确保Excel或WPS已打开
• 准备包含发货明细和账单明细的Excel文件
• 确保文件格式为.xlsx或.xls

第二步：选择工作簿
• 在"发货明细"区域选择包含发货信息的工作簿
• 在"账单明细"区域选择包含账单信息的工作簿
• 工具会自动检测打开的文件

第三步：选择工作表
• 发货明细：选择包含发货信息的工作表（如"发货明细"、"发货"等）
• 账单明细：选择包含账单信息的工作表（如"账单明细"、"账单"等）
• 工具会智能推荐最匹配的工作表

第四步：配置列映射
• 运单号列：选择包含快递单号、运单号等的列
• 商品编码列：选择包含商品编码、SKU等的列
• 商品名称列：选择包含商品名称、品名等的列
• 工具会自动识别并推荐最合适的列

第五步：设置任务选项
• 分隔符：设置多个商品信息之间的分隔符（默认：、）
• 去重：选择是否去除重复的商品信息
• 排序：选择是否对结果进行排序

第六步：预览和开始
• 查看"写入效果预览"确认结果
• 点击"开始任务"执行匹配
• 等待任务完成

注意事项：
• 首次使用建议先在小数据上测试
• 确保Excel文件没有被其他程序占用
• 大文件处理可能需要较长时间，请耐心等待";

            quickStartTab.Controls.Add(richTextBox);
        }

        private void CreateDetailedGuideTab()
        {
            detailedGuideTab = new TabPage("详细指南");
            detailedGuideTab.BackColor = Color.White;

            var richTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("微软雅黑", 10F)
            };

            richTextBox.Text = @"详细使用指南

一、数据准备要求

1. 发货明细数据要求：
   • 必须包含运单号列（快递单号、运单号、邮件号等）
   • 必须包含商品编码列（商品编码、SKU、产品编号等）
   • 必须包含商品名称列（商品名称、品名、产品名称等）
   • 数据应该从第2行开始（第1行作为列标题）

2. 账单明细数据要求：
   • 必须包含运单号列（与发货明细的运单号对应）
   • 可以包含其他需要填充的列
   • 数据应该从第2行开始（第1行作为列标题）

二、智能匹配功能

1. 工作表智能匹配：
   • 工具会自动识别包含"发货明细"、"发货"等关键字的工作表
   • 优先匹配完全匹配的工作表名称
   • 支持模糊匹配，提高识别准确率

2. 列智能匹配：
   • 运单号列：自动识别包含"快递单号"、"运单号"、"邮件号"等关键字的列
   • 商品编码列：自动识别包含"商品编码"、"SKU"、"产品编号"等关键字的列
   • 商品名称列：自动识别包含"商品名称"、"品名"、"产品名称"等关键字的列

三、高级功能

1. 缓存机制：
   • 工具会自动缓存已读取的文件信息
   • 提高重复操作的处理速度
   • 支持手动清理缓存

2. 异步处理：
   • 大文件处理采用异步方式
   • 不会阻塞用户界面
   • 支持进度显示和取消操作

3. 错误处理：
   • 详细的错误提示信息
   • 自动记录操作日志
   • 支持错误恢复

四、性能优化建议

1. 文件大小：
   • 建议单个文件不超过500MB
   • 大文件建议分批处理
   • 关闭不必要的Excel功能

2. 内存管理：
   • 定期清理缓存
   • 避免同时打开过多文件
   • 及时关闭不需要的工作簿

3. 系统资源：
   • 确保有足够的磁盘空间
   • 关闭其他占用内存的程序
   • 使用SSD硬盘提高I/O性能";

            detailedGuideTab.Controls.Add(richTextBox);
        }

        private void CreateFaqTab()
        {
            faqTab = new TabPage("常见问题");
            faqTab.BackColor = Color.White;

            var richTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("微软雅黑", 10F)
            };

            richTextBox.Text = @"常见问题解答

Q1: 工具无法检测到Excel文件怎么办？
A1: 
• 确保Excel或WPS已经打开
• 检查文件是否被其他程序占用
• 尝试重新打开Excel文件
• 检查文件格式是否为.xlsx或.xls

Q2: 智能匹配不准确怎么办？
A2: 
• 检查工作表名称是否包含相关关键字
• 检查列标题是否清晰明确
• 可以手动选择正确的工作表和列
• 在设置中调整智能匹配规则

Q3: 处理大文件时程序卡死怎么办？
A3: 
• 检查文件大小，建议不超过500MB
• 确保有足够的内存空间
• 关闭其他不必要的程序
• 使用异步处理模式

Q4: 匹配结果不完整怎么办？
A4: 
• 检查运单号是否完全一致
• 确认数据格式是否统一
• 检查是否有隐藏字符或空格
• 验证数据完整性

Q5: 如何提高处理速度？
A5: 
• 使用SSD硬盘
• 关闭Excel的自动计算功能
• 减少同时打开的文件数量
• 定期清理缓存

Q6: 程序出现错误怎么办？
A6: 
• 查看错误日志文件
• 重启程序
• 检查数据格式是否正确
• 联系技术支持

Q7: 如何备份配置？
A7: 
• 配置文件保存在用户数据目录
• 可以手动复制配置文件
• 支持配置导入导出功能
• 建议定期备份重要配置

Q8: 支持哪些Excel版本？
A8: 
• Excel 2007及以上版本
• WPS Office
• 支持.xlsx和.xls格式
• 建议使用最新版本";

            faqTab.Controls.Add(richTextBox);
        }

        private void CreateAboutTab()
        {
            aboutTab = new TabPage("关于");
            aboutTab.BackColor = Color.White;

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            // 标题
            var titleLabel = new Label
            {
                Text = Constants.AppName,
                Font = new Font("微软雅黑", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204),
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 40),
                Location = new Point(200, 50)
            };

            // 版本信息
            var versionLabel = new Label
            {
                Text = $"版本: {Constants.AppVersion}",
                Font = new Font("微软雅黑", 12F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 30),
                Location = new Point(200, 100)
            };

            // 公司信息
            var companyLabel = new Label
            {
                Text = $"公司: {Constants.AppCompany}",
                Font = new Font("微软雅黑", 12F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 30),
                Location = new Point(200, 140)
            };

            // 功能特性
            var featuresLabel = new Label
            {
                Text = @"主要功能特性：

• 智能工作表识别和匹配
• 智能列识别和映射
• 高性能缓存机制
• 异步处理大文件
• 详细的日志记录
• 现代化的用户界面
• 完善的错误处理
• 支持多种Excel格式",
                Font = new Font("微软雅黑", 10F),
                ForeColor = Color.Black,
                Size = new Size(400, 200),
                Location = new Point(200, 200)
            };

            // 版权信息
            var copyrightLabel = new Label
            {
                Text = "© 2024 YY Tools. 保留所有权利。",
                Font = new Font("微软雅黑", 9F),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 30),
                Location = new Point(200, 420)
            };

            // 添加控件
            panel.Controls.Add(titleLabel);
            panel.Controls.Add(versionLabel);
            panel.Controls.Add(companyLabel);
            panel.Controls.Add(featuresLabel);
            panel.Controls.Add(copyrightLabel);

            aboutTab.Controls.Add(panel);
        }

        private void ApplyUIEnhancement()
        {
            try
            {
                // 应用UI美化
                UIEnhancer.EnhanceForm(this);
                UIEnhancer.EnhanceAllControls(this);

                // 设置标签页样式
                tabControl.Font = new Font("微软雅黑", 9F);
                tabControl.BackColor = Color.White;
            }
            catch (Exception ex)
            {
                Logger.LogError("应用UI美化失败", ex);
            }
        }

        /// <summary>
        /// 显示用户引导窗体
        /// </summary>
        public static void ShowUserGuide(IWin32Window owner = null)
        {
            try
            {
                var form = new UserGuideForm();
                form.ShowDialog(owner);
            }
            catch (Exception ex)
            {
                Logger.LogError("显示用户引导窗体失败", ex);
                MessageBox.Show("显示用户引导失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}