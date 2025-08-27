// --- 文件 4: ExcelHelper.cs ---
using System;
using System.Linq;

namespace YYTools
{
    public static class ExcelHelper
    {
        public static string GetColumnLetter(int columnNumber)
        {
            if (columnNumber <= 0)
                throw new ArgumentException("列号必须大于0");

            string columnLetter = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                columnLetter = (char)('A' + columnNumber % 26) + columnLetter;
                columnNumber /= 26;
            }
            return columnLetter;
        }

        public static int GetColumnNumber(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter))
                throw new ArgumentException("列字母不能为空");

            columnLetter = columnLetter.ToUpper();
            int columnNumber = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                char letter = columnLetter[i];
                if (letter < 'A' || letter > 'Z')
                    throw new ArgumentException($"无效的列字母：{letter}");

                columnNumber = columnNumber * 26 + (letter - 'A' + 1);
            }

            return columnNumber;
        }
        
        public static bool IsValidColumnLetter(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter))
                return false;
            return columnLetter.ToUpper().All(c => c >= 'A' && c <= 'Z');
        }
    }
}