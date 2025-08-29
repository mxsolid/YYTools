namespace YYTools
{
    /// <summary>
    /// 应用程序常量定义
    /// </summary>
    public static class Constants
    {
        // 应用程序基本信息
        public const string AppName = "YY 运单匹配工具";
        public const string AppVersion = "v3.2 (性能优化版)";
        public const string AppVersionHash = "2024-12-19-8F7E2D1A"; // 唯一版本哈希值
        public const string AppCompany = "YY Tools";
        
        // 配置文件相关
        public const string ConfigFileName = "settings.ini";
        public const string LogFolderName = "Logs";
        public const string CacheFolderName = "Cache";
        
        // UI文本常量
        public const string PreviewPrefixProductCode = "商品: ";
        public const string PreviewPrefixProductName = "品名: ";
        public const string LoadingText = "正在处理...";
        public const string ProcessingText = "处理中...";
        public const string CompletedText = "完成";
        public const string ErrorText = "错误";
        public const string WarningText = "警告";
        
        // 首次运行引导
        public const string FirstRunMessageTitle = "欢迎使用！";
        public const string FirstRunMessage = "欢迎使用 YY 运单匹配工具！\n\n" +
                                              "基本操作步骤：\n" +
                                              "1. 在\"发货明细\"和\"账单明细\"中分别选择对应的工作簿和工作表。\n" +
                                              "2. 工具会自动智能选择关键列（运单号、商品编码等），您也可以手动修改。\n" +
                                              "3. 在\"任务选项\"中配置拼接方式。\n" +
                                              "4. 查看\"写入效果预览\"确认无误后，点击\"开始任务\"。\n\n" +
                                              "遇到问题可以从\"帮助\"菜单查看日志。";
        
        // 状态消息
        public const string StatusNoWorkbooks = "未检测到打开的Excel/WPS文件。请打开文件或从菜单栏选择文件。";
        public const string StatusWorkbooksLoaded = "已加载 {0} 个工作簿。请配置并开始任务。";
        public const string StatusProcessing = "正在处理...";
        public const string StatusCompleted = "任务完成！";
        public const string StatusError = "发生错误：{0}";
        
        // 错误消息
        public const string ErrorNoWorkbooks = "未找到打开的工作簿";
        public const string ErrorNoSheets = "未找到工作表";
        public const string ErrorNoColumns = "未找到列信息";
        public const string ErrorInvalidColumn = "无效的列选择";
        public const string ErrorFileAccess = "无法访问文件：{0}";
        public const string ErrorExcelOperation = "Excel操作失败：{0}";
        public const string ErrorMemoryOverflow = "内存不足，请关闭其他程序后重试";
        
        // 成功消息
        public const string SuccessTaskCompleted = "任务完成！处理 {0} 行，匹配 {1} 个运单，耗时 {2:F2} 秒";
        public const string SuccessSettingsSaved = "设置已保存";
        public const string SuccessCacheCleared = "缓存已清理";
        
        // 工作表智能匹配关键字
        public static readonly string[] ShippingSheetKeywords = { "发货明细", "发货" };
        public static readonly string[] BillSheetKeywords = { "账单明细", "账单" };
        
        // 列类型关键字
        public static readonly string[] TrackColumnKeywords = { "快递单号", "运单号", "邮件号", "物流单号", "快递号", "运单", "快递", "物流", "单号", "tracking" };
        public static readonly string[] ProductColumnKeywords = { "商品编码", "商品编号", "产品编码", "sku", "编码", "商品" };
        public static readonly string[] NameColumnKeywords = { "商品名称", "品名", "名称" };
        
        // 性能配置
        public const int DefaultBatchSize = 1000;
        public const int DefaultMaxPreviewRows = 20; // 默认预览行数改为20
        public const int DefaultMaxThreads = 4;
        public const int DefaultCacheExpirationMinutes = 30;
        
        // 写入预览解析行数选项
        public static readonly int[] PreviewRowOptions = { 5, 10, 20, 50, 100 };
        public const int DefaultPreviewRows = 20; // 默认选择20行
        
        // 文件大小限制 (MB)
        public const long MaxFileSizeMB = 500;
        public const long LargeFileSizeMB = 100;
        
        // 进度条配置
        public const int ProgressBarMinValue = 0;
        public const int ProgressBarMaxValue = 100;
        public const int ProgressBarStep = 1;
        
        // 日志配置
        public const int MaxLogFiles = 30;
        public const int MaxLogFileSizeMB = 10;
        
        // 缓存配置
        public const int MaxCachedWorkbooks = 20;
        public const int MaxCachedWorksheets = 50;
        public const int MaxCachedColumns = 100;
        
        // UI配置
        public const int DefaultFontSize = 9;
        public const int MinFontSize = 8;
        public const int MaxFontSize = 16;
        public const int DefaultFormWidth = 800;
        public const int DefaultFormHeight = 600;
        
        // 动画配置
        public const int LoadingAnimationInterval = 100;
        public const int ProgressUpdateInterval = 50;
        
        // 分隔符选项
        public static readonly string[] DelimiterOptions = { "、", ",", ";", "|", " ", "换行" };
        
        // 排序选项
        public static readonly string[] SortOptions = { "无排序", "升序", "降序" };
        
        // 用户引导文本常量
        public const string QuickStartGuide = "快速开始指南\n\n" +
            "第一步：准备工作\n" +
            "• 确保Excel或WPS已打开\n" +
            "• 准备包含发货明细和账单明细的Excel文件\n" +
            "• 确保文件格式为.xlsx或.xls\n\n" +
            "第二步：选择工作簿\n" +
            "• 在\"发货明细\"区域选择包含发货信息的工作簿\n" +
            "• 在\"账单明细\"区域选择包含账单信息的工作簿\n" +
            "• 工具会自动检测打开的文件\n\n" +
            "第三步：选择工作表\n" +
            "• 发货明细：选择包含发货信息的工作表（如\"发货明细\"、\"发货\"等）\n" +
            "• 账单明细：选择包含账单信息的工作表（如\"账单明细\"、\"账单\"等）\n" +
            "• 工具会智能推荐最匹配的工作表\n\n" +
            "第四步：配置列映射\n" +
            "• 运单号列：选择包含快递单号、运单号等的列\n" +
            "• 商品编码列：选择包含商品编码、SKU等的列\n" +
            "• 商品名称列：选择包含商品名称、品名等的列\n" +
            "• 工具会自动识别并推荐最合适的列\n\n" +
            "第五步：设置任务选项\n" +
            "• 分隔符：设置多个商品信息之间的分隔符（默认：、）\n" +
            "• 去重：选择是否去除重复的商品信息\n" +
            "• 排序：选择是否对结果进行排序\n" +
            "• 预览行数：设置写入预览时解析的行数（默认：20行）\n\n" +
            "第六步：预览和开始\n" +
            "• 查看\"写入效果预览\"确认结果\n" +
            "• 点击\"开始任务\"执行匹配\n" +
            "• 等待任务完成\n\n" +
            "注意事项：\n" +
            "• 首次使用建议先在小数据上测试\n" +
            "• 确保Excel文件没有被其他程序占用\n" +
            "• 大文件处理可能需要较长时间，请耐心等待";
            
        public const string DetailedGuide = "详细使用指南\n\n" +
            "一、数据准备要求\n\n" +
            "1. 发货明细数据要求：\n" +
            "   • 必须包含运单号列（快递单号、运单号、邮件号等）\n" +
            "   • 必须包含商品编码列（商品编码、SKU、产品编号等）\n" +
            "   • 必须包含商品名称列（商品名称、品名、产品名称等）\n" +
            "   • 数据应该从第2行开始（第1行作为列标题）\n\n" +
            "2. 账单明细数据要求：\n" +
            "   • 必须包含运单号列（与发货明细的运单号对应）\n" +
            "   • 可以包含其他需要填充的列\n" +
            "   • 数据应该从第2行开始（第1行作为列标题）\n\n" +
            "二、智能匹配功能\n\n" +
            "1. 工作表智能匹配：\n" +
            "   • 工具会自动识别包含\"发货明细\"、\"发货\"等关键字的工作表\n" +
            "   • 优先匹配完全匹配的工作表名称\n" +
            "   • 支持模糊匹配，提高识别准确率\n\n" +
            "2. 列智能匹配：\n" +
            "   • 运单号列：自动识别包含\"快递单号\"、\"运单号\"、\"邮件号\"等关键字的列\n" +
            "   • 商品编码列：自动识别包含\"商品编码\"、\"SKU\"、\"产品编号\"等关键字的列\n" +
            "   • 商品名称列：自动识别包含\"商品名称\"、\"品名\"、\"产品名称\"等关键字的列\n\n" +
            "三、高级功能\n\n" +
            "1. 缓存机制：\n" +
            "   • 工具会自动缓存已读取的文件信息\n" +
            "   • 提高重复操作的处理速度\n" +
            "   • 支持手动清理缓存\n\n" +
            "2. 异步处理：\n" +
            "   • 大文件处理采用异步方式\n" +
            "   • 不会阻塞用户界面\n" +
            "   • 支持进度显示和取消操作\n\n" +
            "3. 错误处理：\n" +
            "   • 详细的错误提示信息\n" +
            "   • 自动记录操作日志\n" +
            "   • 支持错误恢复\n\n" +
            "四、性能优化建议\n\n" +
            "1. 文件大小：\n" +
            "   • 建议单个文件不超过500MB\n" +
            "   • 大文件建议分批处理\n" +
            "   • 关闭不必要的Excel功能\n\n" +
            "2. 内存管理：\n" +
            "   • 定期清理缓存\n" +
            "   • 避免同时打开过多文件\n" +
            "   • 及时关闭不需要的工作簿\n\n" +
            "3. 系统资源：\n" +
            "   • 确保有足够的磁盘空间\n" +
            "   • 关闭其他占用内存的程序\n" +
            "   • 使用SSD硬盘提高I/O性能";
            
        public const string FrequentlyAskedQuestions = "常见问题解答\n\n" +
            "Q1: 工具无法检测到Excel文件怎么办？\n" +
            "A1: \n" +
            "• 确保Excel或WPS已经打开\n" +
            "• 检查文件是否被其他程序占用\n" +
            "• 尝试重新打开Excel文件\n" +
            "• 检查文件格式是否为.xlsx或.xls\n\n" +
            "Q2: 智能匹配不准确怎么办？\n" +
            "A2: \n" +
            "• 检查工作表名称是否包含相关关键字\n" +
            "• 检查列标题是否清晰明确\n" +
            "• 可以手动选择正确的工作表和列\n" +
            "• 在设置中调整智能匹配规则\n\n" +
            "Q3: 处理大文件时程序卡死怎么办？\n" +
            "A3: \n" +
            "• 检查文件大小，建议不超过500MB\n" +
            "• 确保有足够的内存空间\n" +
            "• 关闭其他不必要的程序\n" +
            "• 使用异步处理模式\n\n" +
            "Q4: 匹配结果不完整怎么办？\n" +
            "A4: \n" +
            "• 检查运单号是否完全一致\n" +
            "• 确认数据格式是否统一\n" +
            "• 检查是否有隐藏字符或空格\n" +
            "• 验证数据完整性\n\n" +
            "Q5: 如何提高处理速度？\n" +
            "A5: \n" +
            "• 使用SSD硬盘\n" +
            "• 关闭Excel的自动计算功能\n" +
            "• 减少同时打开的文件数量\n" +
            "• 定期清理缓存\n\n" +
            "Q6: 程序出现错误怎么办？\n" +
            "A6: \n" +
            "• 查看错误日志文件\n" +
            "• 重启程序\n" +
            "• 检查数据格式是否正确\n" +
            "• 联系技术支持\n\n" +
            "Q7: 如何备份配置？\n" +
            "A7: \n" +
            "• 配置文件保存在用户数据目录\n" +
            "• 可以手动复制配置文件\n" +
            "• 支持配置导入导出功能\n" +
            "• 建议定期备份重要配置\n\n" +
            "Q8: 支持哪些Excel版本？\n" +
            "A8: \n" +
            "• Excel 2007及以上版本\n" +
            "• WPS Office\n" +
            "• 支持.xlsx和.xls格式\n" +
            "• 建议使用最新版本";
            
        public const string AppFeatures = "主要功能特性：\n\n" +
            "• 智能工作表识别和匹配\n" +
            "• 智能列识别和映射\n" +
            "• 高性能缓存机制\n" +
            "• 异步处理大文件\n" +
            "• 详细的日志记录\n" +
            "• 现代化的用户界面\n" +
            "• 完善的错误处理\n" +
            "• 支持多种Excel格式\n" +
            "• 多线程并行处理\n" +
            "• 写入预览行数可配置";
    }
}