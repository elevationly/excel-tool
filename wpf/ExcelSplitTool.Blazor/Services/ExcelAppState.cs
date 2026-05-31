using System.Data;
using System.IO;
using AntDesign;
using ExcelSplitTool.Core.Models;
using ExcelSplitTool.Core.Services;
using Microsoft.AspNetCore.Components;

namespace ExcelSplitTool.Blazor.Services;

public sealed class ExcelAppState
{
    private readonly IFileDialogService _files;
    private readonly IMessageService _message;

    public ExcelAppState(IFileDialogService files, IMessageService message)
    {
        _files = files;
        _message = message;
    }

    public string? FilePath { get; private set; }
    public DataTable? Data { get; private set; }
    public List<List<object?>>? HeaderRows { get; private set; }
    public List<string> SheetNames { get; } = new();
    public string? SelectedSheet { get; set; }
    public List<string> SplitColumns { get; } = new();
    public List<FilterDefinition> Filters { get; } = new();
    public List<ColumnTypeRow> TypeRows { get; } = new();
    public Dictionary<string, ColumnKind> AutoKinds { get; } = new();
    public Dictionary<string, string> CorrectedLabels { get; } = new();
    public SplitExportResult? LastResult { get; private set; }

    public string FileStatus { get; private set; } = "未加载文件";
    public string SheetStatus { get; private set; } = "无工作表";
    public string StatusText { get; private set; } = "请选择 Excel 文件";
    public string OutputHint { get; private set; } = "（选择文件后显示）";
    public string OutputMode { get; set; } = ExcelProcessor.OutputModeFiles;
    public string TypeTab { get; set; } = "all";
    public bool IsRunning { get; private set; }
    public int ProgressPercent { get; private set; }
    public bool ShowProgress { get; private set; }
    public bool CanOpenOutput => LastResult != null;
    public string? AlertMessage { get; private set; }
    public AlertType AlertType { get; private set; } = AlertType.Warning;

    public string? SelectedSplitColumn { get; set; }
    public string? SelectedFilterColumn { get; set; }

    public event Action? Changed;

    public IEnumerable<ColumnTypeRow> FilteredTypeRows
    {
        get
        {
            IEnumerable<ColumnTypeRow> rows = TypeRows;
            return TypeTab switch
            {
                "text" => rows.Where(r => GetColKind(r.ColumnName) == ColumnKind.Text),
                "date" => rows.Where(r => GetColKind(r.ColumnName) == ColumnKind.Date),
                "number" => rows.Where(r => GetColKind(r.ColumnName) == ColumnKind.Number),
                _ => rows
            };
        }
    }

    public List<string> ColumnNames =>
        Data?.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList() ?? new();

    public void Notify() => Changed?.Invoke();

    public async Task PickFileAsync()
    {
        var path = _files.PickExcelFile();
        if (path == null) return;
        try
        {
            FilePath = path;
            SheetNames.Clear();
            SheetNames.AddRange(ExcelProcessor.ListSheetNames(path));
            SelectedSheet = SheetNames.FirstOrDefault();
            FileStatus = Path.GetFileName(path);
            SplitColumns.Clear();
            Filters.Clear();
            StatusText = "已选择文件，正在加载…";
            Notify();
            if (SelectedSheet != null)
                await LoadSheetAsync();
        }
        catch (Exception ex)
        {
            await _message.ErrorAsync(FormatError(ex), 4);
        }
    }

    public async Task ReloadSheetAsync()
    {
        if (string.IsNullOrEmpty(FilePath) || string.IsNullOrEmpty(SelectedSheet)) return;
        await LoadSheetAsync();
    }

    private async Task LoadSheetAsync()
    {
        try
        {
            await Task.Run(() =>
            {
                var sheet = ExcelProcessor.ReadExcelWithHeaders(FilePath!, SelectedSheet!);
                Data = sheet.Data;
                HeaderRows = sheet.HeaderRows;
                AutoKinds.Clear();
                CorrectedLabels.Clear();
                foreach (var kv in ExcelProcessor.DetectAllColumnKinds(Data))
                {
                    AutoKinds[kv.Key] = kv.Value;
                    CorrectedLabels[kv.Key] = ColumnKindLabels.ToLabel(kv.Value);
                }
            });
            SplitColumns.Clear();
            Filters.Clear();
            RebuildTypeRows();
            SheetStatus = $"{SelectedSheet} · {Data!.Rows.Count} 行";
            StatusText = $"已加载：{SelectedSheet}（{Data.Rows.Count} 行）";
            RefreshOutputHint();
            Notify();
        }
        catch (Exception ex)
        {
            await _message.ErrorAsync(FormatError(ex), 4);
        }
    }

    public async Task AddSplitColumnAsync()
    {
        if (string.IsNullOrEmpty(SelectedSplitColumn))
        {
            await ShowAlertAsync("请先选择要添加的拆分列");
            return;
        }
        if (SplitColumns.Contains(SelectedSplitColumn))
        {
            await ShowAlertAsync("该列已在拆分列表中");
            return;
        }
        SplitColumns.Add(SelectedSplitColumn);
        ClearAlert();
        Notify();
    }

    public void RemoveSplitColumn(string? col)
    {
        if (!string.IsNullOrEmpty(col) && SplitColumns.Remove(col)) { Notify(); return; }
        RemoveLastSplitColumn();
    }

    public void RemoveLastSplitColumn()
    {
        if (SplitColumns.Count > 0) SplitColumns.RemoveAt(SplitColumns.Count - 1);
        Notify();
    }

    public void AddFilter()
    {
        if (string.IsNullOrEmpty(SelectedFilterColumn)) return;
        if (Filters.Any(f => f.Column == SelectedFilterColumn)) return;
        Filters.Add(new FilterDefinition
        {
            Column = SelectedFilterColumn,
            Kind = GetColKind(SelectedFilterColumn)
        });
        Notify();
    }

    public void RemoveFilter(FilterDefinition f)
    {
        Filters.Remove(f);
        Notify();
    }

    public void RemoveLastFilter()
    {
        if (Filters.Count > 0) Filters.RemoveAt(Filters.Count - 1);
        Notify();
    }

    public void ResetKinds()
    {
        if (Data == null) return;
        CorrectedLabels.Clear();
        foreach (var kv in AutoKinds)
            CorrectedLabels[kv.Key] = ColumnKindLabels.ToLabel(kv.Value);
        RebuildTypeRows();
        Notify();
    }

    public void SetTypeTab(string tab)
    {
        TypeTab = tab;
        Notify();
    }

    public void OnKindChanged(ColumnTypeRow row, string label)
    {
        row.CorrectedLabel = label;
        CorrectedLabels[row.ColumnName] = label;
        Notify();
    }

    public void RefreshOutputHint()
    {
        if (string.IsNullOrEmpty(FilePath))
        {
            OutputHint = "（选择文件后显示）";
            return;
        }
        OutputHint = OutputMode == ExcelProcessor.OutputModeWorkbook
            ? $"将生成：{ExcelProcessor.MakeTimestampedOutputFile(FilePath)}"
            : $"将生成文件夹：{ExcelProcessor.MakeTimestampedOutputDir(FilePath)}";
    }

    public async Task RunSplitAsync()
    {
        if (Data == null || string.IsNullOrEmpty(FilePath))
        {
            await ShowAlertAsync("请先选择并加载 Excel 文件");
            return;
        }
        if (SplitColumns.Count == 0)
        {
            await ShowAlertAsync("请至少添加一列用于拆分");
            return;
        }

        ClearAlert();

        foreach (var row in TypeRows)
            CorrectedLabels[row.ColumnName] = row.CorrectedLabel;

        IsRunning = true;
        ShowProgress = true;
        ProgressPercent = 0;
        StatusText = "正在拆分…";
        Notify();

        var progress = new Progress<(int current, int total)>(p =>
        {
            if (p.total <= 0) return;
            ProgressPercent = (int)Math.Min(100, p.current * 100.0 / p.total);
            StatusText = $"正在拆分… ({p.current}/{p.total})";
            Notify();
        });

        try
        {
            var sourceFile = FilePath;
            var mode = OutputMode;
            var outputPath = mode == ExcelProcessor.OutputModeWorkbook
                ? ExcelProcessor.MakeTimestampedOutputFile(sourceFile)
                : ExcelProcessor.MakeTimestampedOutputDir(sourceFile);
            var kinds = GetColKindsDict();
            var filtered = ExcelProcessor.ApplyFilters(Data, Filters);
            var result = await Task.Run(() =>
                ExcelProcessor.SplitAndExport(
                    filtered, SplitColumns, outputPath, HeaderRows, kinds, mode, progress));
            LastResult = result;
            ProgressPercent = 100;
            StatusText = $"完成：{result.Count} 项";
            Notify();

            var msg = mode == ExcelProcessor.OutputModeWorkbook
                ? $"已写入 {result.Count} 个工作表"
                : $"已生成 {result.Count} 个文件";
            await _message.SuccessAsync($"{msg}，输出：{result.OutputPath}", 5);
        }
        catch (Exception ex)
        {
            var msg = FormatError(ex);
            await ShowAlertAsync(msg, AlertType.Error);
            StatusText = "拆分失败";
        }
        finally
        {
            IsRunning = false;
            ShowProgress = false;
            Notify();
        }
    }

    public void OpenOutput()
    {
        if (LastResult != null) _files.OpenPath(LastResult.OutputPath);
    }

    private void RebuildTypeRows()
    {
        TypeRows.Clear();
        if (Data == null) return;
        foreach (DataColumn col in Data.Columns)
        {
            var name = col.ColumnName;
            var auto = AutoKinds.TryGetValue(name, out var k) ? k : ColumnKind.Text;
            var autoLabel = ColumnKindLabels.ToLabel(auto);
            TypeRows.Add(new ColumnTypeRow
            {
                ColumnName = name,
                AutoLabel = autoLabel,
                CorrectedLabel = CorrectedLabels.TryGetValue(name, out var c) ? c : autoLabel
            });
        }
    }

    private ColumnKind GetColKind(string col)
    {
        if (CorrectedLabels.TryGetValue(col, out var label))
            return ColumnKindLabels.FromLabel(label);
        return AutoKinds.TryGetValue(col, out var k) ? k : ColumnKind.Text;
    }

    private Dictionary<string, ColumnKind> GetColKindsDict() =>
        Data?.Columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => GetColKind(c.ColumnName))
        ?? new Dictionary<string, ColumnKind>();

    public void ClearAlert()
    {
        AlertMessage = null;
        Notify();
    }

    private async Task ShowAlertAsync(string message, AlertType type = AlertType.Warning)
    {
        AlertMessage = message;
        AlertType = type;
        StatusText = message;
        Notify();
        try
        {
            if (type == AlertType.Error)
                await _message.ErrorAsync(message);
            else
                await _message.WarningAsync(message);
        }
        catch
        {
            // WebView 中 Message 可能不可用，以页面 Alert 为准
        }
    }

    private static string FormatError(Exception ex)
    {
        if (ex is TypeInitializationException tie && tie.InnerException != null)
            ex = tie.InnerException;
        if (ex is ArgumentException) return ex.Message;
        var msg = ex.Message;
        if (ex.InnerException != null && !msg.Contains(ex.InnerException.Message))
            msg += " — " + ex.InnerException.Message;
        return msg;
    }
}
