namespace YYTools
{
    public static class Constants
    {
        public const string AppName = "YY 运单匹配工具";
        public const string AppVersion = "v3.0 (重构版)";

        public const string PreviewPrefixProductCode = "商品: ";
        public const string PreviewPrefixProductName = "品名: ";

        public const string FirstRunMessageTitle = "欢迎使用！";
        public const string FirstRunMessage = "欢迎使用 YY 运单匹配工具！\n\n" +
                                              "基本操作步骤：\n" +
                                              "1. 在“发货明细”和“账单明细”中分别选择对应的工作簿和工作表。\n" +
                                              "2. 工具会自动智能选择关键列（运单号、商品编码等），您也可以手动修改。\n" +
                                              "3. 在“任务选项”中配置拼接方式。\n" +
                                              "4. 查看“写入效果预览”确认无误后，点击“开始任务”。\n\n" +
                                              "遇到问题可以从“帮助”菜单查看日志。";
    }
}