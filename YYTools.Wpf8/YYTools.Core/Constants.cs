namespace YYTools
{
    /// <summary>
    /// 应用程序常量定义（从旧版移植）
    /// </summary>
    public static class Constants
    {
        public const string AppName = "YY 运单匹配工具";
        public const string AppVersion = "v3.2 (性能优化版)";
        public const string AppVersionHash = "2024-12-19-8F7E2D1A";
        public const string AppCompany = "YY Tools";

        public const string ConfigFileName = "settings.ini";
        public const string LogFolderName = "Logs";
        public const string CacheFolderName = "Cache";

        public const string PreviewPrefixProductCode = "商品: ";
        public const string PreviewPrefixProductName = "品名: ";
        public const string LoadingText = "正在处理...";
        public const string ProcessingText = "处理中...";
        public const string CompletedText = "完成";
        public const string ErrorText = "错误";
        public const string WarningText = "警告";

        public const string FirstRunMessageTitle = "欢迎使用！";
        public const string FirstRunMessage = "欢迎使用 YY 运单匹配工具！\n\n基本操作步骤：\n1. 在\"发货明细\"和\"账单明细\"中分别选择对应的工作簿和工作表。\n2. 工具会自动智能选择关键列（运单号、商品编码等），您也可以手动修改。\n3. 在\"任务选项\"中配置拼接方式。\n4. 查看\"写入效果预览\"确认无误后，点击\"开始任务\"。\n\n遇到问题可以从\"帮助\"菜单查看日志。";

        public const string StatusNoWorkbooks = "未检测到打开的Excel/WPS文件。请打开文件或从菜单栏选择文件。";
        public const string StatusWorkbooksLoaded = "已加载 {0} 个工作簿。请配置并开始任务。";
        public const string StatusProcessing = "正在处理...";
        public const string StatusCompleted = "任务完成！";
        public const string StatusError = "发生错误：{0}";

        public const string ErrorNoWorkbooks = "未找到打开的工作簿";
        public const string ErrorNoSheets = "未找到工作表";
        public const string ErrorNoColumns = "未找到列信息";
        public const string ErrorInvalidColumn = "无效的列选择";
        public const string ErrorFileAccess = "无法访问文件：{0}";
        public const string ErrorExcelOperation = "Excel操作失败：{0}";
        public const string ErrorMemoryOverflow = "内存不足，请关闭其他程序后重试";

        public const string SuccessTaskCompleted = "任务完成！处理 {0} 行，匹配 {1} 个运单，耗时 {2:F2} 秒";
        public const string SuccessSettingsSaved = "设置已保存";
        public const string SuccessCacheCleared = "缓存已清理";

        public static readonly string[] ShippingSheetKeywords = { "发货明细", "发货" };
        public static readonly string[] BillSheetKeywords = { "账单明细", "账单" };

        public static readonly string[] TrackColumnKeywords = { "快递单号", "运单号", "邮件号", "物流单号", "快递号", "运单", "快递", "物流", "单号", "tracking" };
        public static readonly string[] ProductColumnKeywords = { "商品编码", "商品编号", "产品编码", "sku", "编码", "商品" };
        public static readonly string[] NameColumnKeywords = { "商品名称", "品名", "名称" };

        public const int DefaultBatchSize = 1000;
        public const int DefaultMaxPreviewRows = 20;
        public const int DefaultMaxThreads = 4;
        public const int DefaultCacheExpirationMinutes = 30;

        public static readonly int[] PreviewRowOptions = { 5, 10, 20, 50, 100 };
        public const int DefaultPreviewRows = 20;
        public const int DefaultPreviewParseRows = 20;

        public const long MaxFileSizeMB = 500;
        public const long LargeFileSizeMB = 100;

        public const int ProgressBarMinValue = 0;
        public const int ProgressBarMaxValue = 100;
        public const int ProgressBarStep = 1;

        public const int MaxLogFiles = 30;
        public const int MaxLogFileSizeMB = 10;

        public const int MaxCachedWorkbooks = 20;
        public const int MaxCachedWorksheets = 50;
        public const int MaxCachedColumns = 100;

        public const int DefaultFontSize = 9;
        public const int MinFontSize = 8;
        public const int MaxFontSize = 16;
        public const int DefaultFormWidth = 800;
        public const int DefaultFormHeight = 600;

        public const int LoadingAnimationInterval = 100;
        public const int ProgressUpdateInterval = 50;

        public static readonly string[] DelimiterOptions = { "、", ",", ";", "|", " ", "换行" };
        public static readonly string[] SortOptions = { "无排序", "升序", "降序" };
    }
}

