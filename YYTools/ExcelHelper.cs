using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YYTools
{
    /// <summary>
    /// Excel帮助类，提供列转换和地址解析等实用功能
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>
        /// 根据列号获取Excel列字母（如：1->A, 27->AA）
        /// </summary>
        /// <param name="columnNumber">列号（从1开始）</param>
        /// <returns>列字母</returns>
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

        /// <summary>
        /// 根据Excel列字母获取列号（如：A->1, AA->27）
        /// </summary>
        /// <param name="columnLetter">列字母</param>
        /// <returns>列号</returns>
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
                    throw new ArgumentException(string.Format("无效的列字母：{0}", letter));
                
                columnNumber = columnNumber * 26 + (letter - 'A' + 1);
            }
            
            return columnNumber;
        }

        /// <summary>
        /// 验证列字母格式是否正确
        /// </summary>
        /// <param name="columnLetter">列字母</param>
        /// <returns>是否有效</returns>
        public static bool IsValidColumnLetter(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter))
                return false;

            try
            {
                GetColumnNumber(columnLetter);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取单元格地址（如：A1, B5, AA10）
        /// </summary>
        /// <param name="columnNumber">列号</param>
        /// <param name="rowNumber">行号</param>
        /// <returns>单元格地址</returns>
        public static string GetCellAddress(int columnNumber, int rowNumber)
        {
            if (columnNumber <= 0 || rowNumber <= 0)
                throw new ArgumentException("行号和列号必须大于0");

            return GetColumnLetter(columnNumber) + rowNumber.ToString();
        }

        /// <summary>
        /// 解析单元格地址，返回列号和行号
        /// </summary>
        /// <param name="cellAddress">单元格地址（如：A1, B5）</param>
        /// <returns>包含列号和行号的结构体</returns>
        public static CellPosition ParseCellAddress(string cellAddress)
        {
            if (string.IsNullOrWhiteSpace(cellAddress))
                throw new ArgumentException("单元格地址不能为空");

            cellAddress = cellAddress.ToUpper().Trim();
            
            // 分离字母和数字部分
            int letterEndIndex = 0;
            while (letterEndIndex < cellAddress.Length && char.IsLetter(cellAddress[letterEndIndex]))
            {
                letterEndIndex++;
            }

            if (letterEndIndex == 0 || letterEndIndex == cellAddress.Length)
                throw new ArgumentException(string.Format("无效的单元格地址：{0}", cellAddress));

            string columnPart = cellAddress.Substring(0, letterEndIndex);
            string rowPart = cellAddress.Substring(letterEndIndex);

            int columnNumber = GetColumnNumber(columnPart);
            
            int rowNumber;
            if (!int.TryParse(rowPart, out rowNumber) || rowNumber <= 0)
                throw new ArgumentException(string.Format("无效的行号：{0}", rowPart));

            return new CellPosition { Column = columnNumber, Row = rowNumber };
        }

        /// <summary>
        /// 获取范围地址（如：A1:B5）
        /// </summary>
        /// <param name="startColumn">起始列号</param>
        /// <param name="startRow">起始行号</param>
        /// <param name="endColumn">结束列号</param>
        /// <param name="endRow">结束行号</param>
        /// <returns>范围地址</returns>
        public static string GetRangeAddress(int startColumn, int startRow, int endColumn, int endRow)
        {
            string startCell = GetCellAddress(startColumn, startRow);
            string endCell = GetCellAddress(endColumn, endRow);
            return string.Format("{0}:{1}", startCell, endCell);
        }

        /// <summary>
        /// 递增列字母（如：A->B, Z->AA）
        /// </summary>
        /// <param name="columnLetter">当前列字母</param>
        /// <param name="increment">递增量，默认为1</param>
        /// <returns>递增后的列字母</returns>
        public static string IncrementColumn(string columnLetter, int increment = 1)
        {
            int columnNumber = GetColumnNumber(columnLetter);
            return GetColumnLetter(columnNumber + increment);
        }

        /// <summary>
        /// 递增行号
        /// </summary>
        /// <param name="rowNumber">当前行号</param>
        /// <param name="increment">递增量，默认为1</param>
        /// <returns>递增后的行号</returns>
        public static int IncrementRow(int rowNumber, int increment = 1)
        {
            return rowNumber + increment;
        }
    }

    /// <summary>
    /// 单元格位置结构体
    /// </summary>
    public struct CellPosition
    {
        public int Column { get; set; }
        public int Row { get; set; }
    }
} 