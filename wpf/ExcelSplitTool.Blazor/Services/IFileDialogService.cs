namespace ExcelSplitTool.Blazor.Services;

public interface IFileDialogService
{
    string? PickExcelFile();
    void OpenPath(string path);
}
