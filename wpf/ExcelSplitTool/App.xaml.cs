using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace ExcelSplitTool;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;

        try
        {
            base.OnStartup(e);
            var win = new MainWindow();
            MainWindow = win;
            win.Show();
        }
        catch (Exception ex)
        {
            ShowFatal(ex);
            Shutdown(1);
        }
    }

    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        ShowFatal(e.Exception);
        e.Handled = true;
        Current.Shutdown(1);
    }

    private static void OnDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
            ShowFatal(ex);
    }

    private static void ShowFatal(Exception ex)
    {
        var msg = ex.ToString();
        try
        {
            var log = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "startup_error.log");
            File.WriteAllText(log, msg);
        }
        catch { /* ignore */ }

        MessageBox.Show(
            "程序启动失败：\n\n" + ex.Message + "\n\n详情已写入 startup_error.log",
            "Excel 表格拆分工具",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }
}
