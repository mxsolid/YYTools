using System.Windows;
using Microsoft.Extensions.DependencyInjection;

namespace YYTools.App
{
	public partial class MainWindow : MahApps.Metro.Controls.MetroWindow
	{
		public MainWindow(ViewModels.MainViewModel viewModel)
		{
			InitializeComponent();
			DataContext = viewModel;
		}
	}
}