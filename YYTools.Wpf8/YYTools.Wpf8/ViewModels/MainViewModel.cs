using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Threading;
using System.Threading.Tasks;

namespace YYTools.Wpf8.ViewModels
{
    /// <summary>
    /// 主视图模型：演示TPL与IProgress用法，后续将承载匹配业务
    /// </summary>
    public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty]
        private string statusText = "就绪";

        [ObservableProperty]
        private int progressValue;

        [RelayCommand]
        private async Task DoLongWorkAsync()
        {
            var progress = new Progress<int>(p => ProgressValue = p);
            StatusText = "处理中...";
            await Task.Run(async () =>
            {
                for (int i = 0; i <= 100; i += 5)
                {
                    ((IProgress<int>)progress).Report(i);
                    await Task.Delay(20);
                }
            });
            StatusText = "完成";
        }
    }
}

