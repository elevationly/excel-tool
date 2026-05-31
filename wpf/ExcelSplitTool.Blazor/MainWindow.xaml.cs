using System.Windows;
using AntDesign;
using ExcelSplitTool.Blazor.Services;
using Microsoft.AspNetCore.Components.WebView.Wpf;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace ExcelSplitTool.Blazor;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        var services = new ServiceCollection();
        services.AddWpfBlazorWebView();
#if DEBUG
        services.AddBlazorWebViewDeveloperTools();
        services.AddLogging(builder => builder.AddDebug());
#endif
        services.AddAntDesign();
        services.AddSingleton<IFileDialogService, WpfFileDialogService>();
        services.AddSingleton<ExcelAppState>();

        blazorWebView.Services = services.BuildServiceProvider();
    }
}
