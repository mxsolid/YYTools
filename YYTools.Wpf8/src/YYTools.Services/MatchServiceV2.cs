using System;
using System.Threading;
using System.Threading.Tasks;

namespace YYTools.Services
{
	/// <summary>
	/// 匹配服务 V2：负责并行匹配与结果汇总，支持 IProgress 上报。
	/// 这里提供骨架，后续可将旧版 MatchService 的核心逻辑移植过来。
	/// </summary>
	public class MatchServiceV2
	{
		public async Task MatchAsync(IProgress<(int,string)>? progress, CancellationToken ct)
		{
			await Task.Run(async () =>
			{
				progress?.Report((10, "准备数据..."));
				await Task.Delay(100, ct);
				progress?.Report((50, "并行匹配中..."));
				await Task.Delay(200, ct);
				progress?.Report((90, "写回结果..."));
				await Task.Delay(100, ct);
				progress?.Report((100, "完成"));
			}, ct);
		}
	}
}