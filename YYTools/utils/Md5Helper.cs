using System;
using System.Security.Cryptography;
using System.Text;

namespace YYTools.Utils
{
    public static class Md5Helper
    {
        /// <summary>
        /// 对字符串生成 MD5 哈希值（32 位十六进制，默认小写）
        /// </summary>
        /// <param name="input">待计算的字符串</param>
        /// <param name="isUpper">是否返回大写格式（默认 false，小写）</param>
        /// <returns>MD5 哈希字符串（32 位）</returns>
        public static string GetStringMd5(string input, bool isUpper = false)
        {
            // 1. 检查输入是否为空
            if (string.IsNullOrEmpty(input))
                throw new ArgumentNullException(nameof(input), "输入字符串不能为空");

            // 2. 将字符串转为字节数组（指定编码为 UTF-8，避免默认编码差异）
            byte[] inputBytes = Encoding.UTF8.GetBytes(input);

            // 3. 创建 MD5 实例，计算哈希值
            using (MD5 md5 = MD5.Create()) // using 自动释放资源，避免内存泄漏
            {
                byte[] hashBytes = md5.ComputeHash(inputBytes); // 核心：计算哈希字节数组

                // 4. 将字节数组转为十六进制字符串
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                {
                    // 格式化为两位十六进制（x2 小写，X2 大写）
                    sb.Append(isUpper ? b.ToString("X2") : b.ToString("x2"));
                }

                return sb.ToString();
            }
        }
    }
}