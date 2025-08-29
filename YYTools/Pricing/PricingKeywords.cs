using System;

namespace YYTools.Pricing
{
    /// <summary>
    /// 价格相关的关键词与分类常量集中定义
    /// 注意：仅提供常量，不包含任何业务逻辑，避免硬编码分散。
    /// </summary>
    public static class PricingKeywords
    {
        /// <summary>
        /// 识别价格/金额类列头的关键词集合（按常见度排列）
        /// </summary>
        public static readonly string[] PriceKeywords = new[]
        {
            "价格", "单价", "金额", "price", "amount", "价钱", "合计", "总价", "总金额"
        };

        /// <summary>
        /// 识别数量类列头的关键词集合（供可能的扩展使用）
        /// </summary>
        public static readonly string[] QuantityKeywords = new[]
        {
            "数量", "件数", "qty", "quantity"
        };
    }
}

