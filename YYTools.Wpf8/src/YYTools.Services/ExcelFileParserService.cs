using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;

namespace YYTools.Services
{
	/// <summary>
	/// 使用 ExcelDataReader 解析本地 Excel 文件（无需安装 Excel）。
	/// </summary>
	public class ExcelFileParserService
	{
		public sealed class ParsedWorkbook
		{
			public List<string> SheetNames { get; init; } = new();
		}

		public async Task<ParsedWorkbook> ParseWorkbookAsync(string filePath, IProgress<(int,string)>? progress, CancellationToken ct)
		{
			return await Task.Run(() =>
			{
				System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
				progress?.Report((10, "打开文件..."));
				var result = new ParsedWorkbook();
				using var stream = File.OpenRead(filePath);
				using var reader = ExcelReaderFactory.CreateReader(stream);
				var dataSet = reader.AsDataSet();
				progress?.Report((60, "解析工作表..."));
				foreach (System.Data.DataTable table in dataSet.Tables)
				{
					result.SheetNames.Add(table.TableName);
				}
				progress?.Report((100, "解析完成"));
				return result;
			}, ct);
		}
	}
}