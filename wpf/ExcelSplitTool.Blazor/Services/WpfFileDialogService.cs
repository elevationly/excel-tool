using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace ExcelSplitTool.Blazor.Services;

public sealed class WpfFileDialogService : IFileDialogService
{
    public string? PickExcelFile()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Excel 文件|*.xlsx;*.xls|所有文件|*.*",
            Title = "选择 Excel 文件"
        };
        return dlg.ShowDialog() == true ? dlg.FileName : null;
    }

    public void OpenPath(string path)
    {
        try
        {
            if (File.Exists(path))
                Process.Start("explorer.exe", $"/select,\"{path}\"");
            else if (Directory.Exists(path))
                Process.Start("explorer.exe", path);
        }
        catch { /* ignore */ }
    }
}
