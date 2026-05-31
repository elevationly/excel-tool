using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelSplitTool.Models;
using ExcelSplitTool.Services;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace ExcelSplitTool;

public partial class MainWindow : Window
{
    private string? _filePath;
    private DataTable? _data;
    private List<List<object?>>? _headerRows;
    private readonly List<string> _splitColumns = new();
    private readonly List<FilterDefinition> _filters = new();
    private readonly Dictionary<string, ColumnKind> _autoKinds = new();
    private readonly Dictionary<string, string> _correctedLabels = new();
    private readonly List<ColumnTypeRow> _typeRows = new();
    private SplitExportResult? _lastResult;
    private string _typeTab = "all";

    public MainWindow()
    {
        InitializeComponent();
        Loaded += OnWindowLoaded;
    }

    private void OnWindowLoaded(object sender, RoutedEventArgs e) => RefreshOutputHint();

    private void BrowseFile_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Excel 文件|*.xlsx;*.xls|所有文件|*.*",
            Title = "选择 Excel 文件"
        };
        if (dlg.ShowDialog() != true) return;
        LoadFile(dlg.FileName);
    }

    private void LoadFile(string path)
    {
        try
        {
            _filePath = path;
            var sheets = ExcelProcessor.ListSheetNames(path);
            CmbSheet.ItemsSource = sheets;
            if (sheets.Count > 0)
                CmbSheet.SelectedIndex = 0;
            TxtFileName.Text = Path.GetFileName(path);
            TxtPathHint.Text = path;
            SetFileStatus(true, Path.GetFileName(path));
            _splitColumns.Clear();
            LstSplitCols.Items.Clear();
            ClearFilters();
            _setStatus("已选择文件，正在加载…");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, FormatUserError(ex), "打开失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string FormatUserError(Exception ex)
    {
        if (ex is TypeInitializationException tie && tie.InnerException != null)
            ex = tie.InnerException;
        var msg = ex.Message;
        if (ex.InnerException != null && !msg.Contains(ex.InnerException.Message))
            msg += "\n\n" + ex.InnerException.Message;
        return msg;
    }

    private void Sheet_Changed(object sender, SelectionChangedEventArgs e) => ReloadSheet();

    private void ReloadSheet_Click(object sender, RoutedEventArgs e) => ReloadSheet();

    private void ReloadSheet()
    {
        if (string.IsNullOrEmpty(_filePath) || CmbSheet.SelectedItem is not string sheet) return;
        try
        {
            var sheetData = ExcelProcessor.ReadExcelWithHeaders(_filePath!, sheet);
            _data = sheetData.Data;
            _headerRows = sheetData.HeaderRows;
            _autoKinds.Clear();
            _correctedLabels.Clear();
            foreach (var kv in ExcelProcessor.DetectAllColumnKinds(_data))
            {
                _autoKinds[kv.Key] = kv.Value;
                _correctedLabels[kv.Key] = ColumnKindLabels.ToLabel(kv.Value);
            }
            RefreshColumnLists();
            RebuildTypeGrid();
            ClearFilters();
            _splitColumns.Clear();
            LstSplitCols.Items.Clear();
            SetSheetStatus(true, $"{sheet} · {_data.Rows.Count} 行");
            _setStatus($"已加载：{sheet}（{_data.Rows.Count} 行）");
            RefreshOutputHint();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, FormatUserError(ex), "加载失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void RefreshColumnLists()
    {
        if (_data == null) return;
        var cols = _data.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
        CmbSplitCol.ItemsSource = cols;
        CmbFilterCol.ItemsSource = cols;
    }

    private ColumnKind GetColKind(string col)
    {
        if (_correctedLabels.TryGetValue(col, out var label))
            return ColumnKindLabels.FromLabel(label);
        return _autoKinds.TryGetValue(col, out var k) ? k : ColumnKind.Text;
    }

    private Dictionary<string, ColumnKind> GetColKindsDict() =>
        _data?.Columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => GetColKind(c.ColumnName))
        ?? new Dictionary<string, ColumnKind>();

    private void AddSplitCol_Click(object sender, RoutedEventArgs e)
    {
        if (CmbSplitCol.SelectedItem is not string col) return;
        if (_splitColumns.Contains(col)) return;
        _splitColumns.Add(col);
        LstSplitCols.Items.Add(col);
    }

    private void RemoveSplitCol_Click(object sender, RoutedEventArgs e)
    {
        if (LstSplitCols.SelectedItem is string col)
        {
            _splitColumns.Remove(col);
            LstSplitCols.Items.Remove(col);
        }
        else if (_splitColumns.Count > 0)
        {
            var last = _splitColumns[_splitColumns.Count - 1];
            _splitColumns.RemoveAt(_splitColumns.Count - 1);
            LstSplitCols.Items.Remove(last);
        }
    }

    private void AddFilter_Click(object sender, RoutedEventArgs e)
    {
        if (CmbFilterCol.SelectedItem is not string col) return;
        if (_filters.Any(f => f.Column == col)) return;
        var f = new FilterDefinition { Column = col, Kind = GetColKind(col) };
        _filters.Add(f);
        RebuildFilterPanel();
    }

    private void RemoveFilter_Click(object sender, RoutedEventArgs e)
    {
        if (_filters.Count == 0) return;
        var last = _filters[_filters.Count - 1];
        _filters.RemoveAt(_filters.Count - 1);
        RebuildFilterPanel();
    }

    private void ClearFilters()
    {
        _filters.Clear();
        FiltersPanel.Children.Clear();
    }

    private void RebuildFilterPanel()
    {
        FiltersPanel.Children.Clear();
        foreach (var f in _filters)
            FiltersPanel.Children.Add(BuildFilterRow(f));
    }

    private UIElement BuildFilterRow(FilterDefinition f)
    {
        var border = new Border
        {
            BorderBrush = (Brush)FindResource("BorderBrush"),
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(0, 6, 0, 6),
            Margin = new Thickness(0, 0, 0, 4)
        };
        var panel = new StackPanel { Orientation = Orientation.Horizontal };
        panel.Children.Add(new TextBlock
        {
            Text = f.Column,
            Width = 100,
            VerticalAlignment = VerticalAlignment.Center,
            FontWeight = FontWeights.SemiBold
        });

        switch (f.Kind)
        {
            case ColumnKind.Date:
                panel.Children.Add(MkLabel("起"));
                var dpStart = new DatePicker { Width = 120, Margin = new Thickness(4, 0, 8, 0) };
                if (f.Start.HasValue) dpStart.SelectedDate = f.Start;
                dpStart.SelectedDateChanged += (_, _) => f.Start = dpStart.SelectedDate;
                panel.Children.Add(dpStart);
                panel.Children.Add(MkLabel("止"));
                var dpEnd = new DatePicker { Width = 120, Margin = new Thickness(4, 0, 8, 0) };
                if (f.End.HasValue) dpEnd.SelectedDate = f.End;
                dpEnd.SelectedDateChanged += (_, _) => f.End = dpEnd.SelectedDate;
                panel.Children.Add(dpEnd);
                break;
            case ColumnKind.Number:
                panel.Children.Add(MkLabel("≥"));
                var txtMin = new System.Windows.Controls.TextBox { Width = 72, Margin = new Thickness(4, 0, 8, 0) };
                if (f.Min.HasValue) txtMin.Text = f.Min.Value.ToString(CultureInfo.InvariantCulture);
                txtMin.TextChanged += (_, _) =>
                {
                    f.Min = double.TryParse(txtMin.Text, out var v) ? v : null;
                };
                panel.Children.Add(txtMin);
                panel.Children.Add(MkLabel("≤"));
                var txtMax = new System.Windows.Controls.TextBox { Width = 72, Margin = new Thickness(4, 0, 8, 0) };
                if (f.Max.HasValue) txtMax.Text = f.Max.Value.ToString(CultureInfo.InvariantCulture);
                txtMax.TextChanged += (_, _) =>
                {
                    f.Max = double.TryParse(txtMax.Text, out var v) ? v : null;
                };
                panel.Children.Add(txtMax);
                break;
            default:
                panel.Children.Add(MkLabel("包含"));
                var txt = new System.Windows.Controls.TextBox { Width = 160, Margin = new Thickness(4, 0, 8, 0) };
                txt.Text = f.Text;
                txt.TextChanged += (_, _) => f.Text = txt.Text;
                panel.Children.Add(txt);
                break;
        }

        var btnRemove = new System.Windows.Controls.Button
        {
            Content = "移除",
            Width = 56,
            Margin = new Thickness(4, 0, 0, 0)
        };
        btnRemove.Click += (_, _) =>
        {
            _filters.Remove(f);
            RebuildFilterPanel();
        };
        panel.Children.Add(btnRemove);
        border.Child = panel;
        return border;
    }

    private static TextBlock MkLabel(string text) => new()
    {
        Text = text,
        VerticalAlignment = VerticalAlignment.Center,
        Foreground = Brushes.Gray
    };

    private void RebuildTypeGrid()
    {
        _typeRows.Clear();
        if (_data == null) return;
        foreach (DataColumn col in _data.Columns)
        {
            var name = col.ColumnName;
            var auto = _autoKinds.TryGetValue(name, out var k) ? k : ColumnKind.Text;
            var autoLabel = ColumnKindLabels.ToLabel(auto);
            _typeRows.Add(new ColumnTypeRow
            {
                ColumnName = name,
                AutoLabel = autoLabel,
                CorrectedLabel = _correctedLabels.TryGetValue(name, out var c) ? c : autoLabel
            });
        }
        ApplyTypeTabFilter();
    }

    private void ApplyTypeTabFilter()
    {
        IEnumerable<ColumnTypeRow> rows = _typeRows;
        rows = _typeTab switch
        {
            "text" => rows.Where(r => GetColKind(r.ColumnName) == ColumnKind.Text),
            "date" => rows.Where(r => GetColKind(r.ColumnName) == ColumnKind.Date),
            "number" => rows.Where(r => GetColKind(r.ColumnName) == ColumnKind.Number),
            _ => rows
        };
        TypeColumnList.ItemsSource = null;
        TypeColumnList.ItemsSource = rows.ToList();
    }

    private void TypeTab_Changed(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded || TabAll == null) return;
        _typeTab = TabText?.IsChecked == true ? "text"
            : TabDate?.IsChecked == true ? "date"
            : TabNumber?.IsChecked == true ? "number"
            : "all";
        ApplyTypeTabFilter();
    }

    private void ResetKinds_Click(object sender, RoutedEventArgs e)
    {
        if (_data == null) return;
        _correctedLabels.Clear();
        foreach (var kv in _autoKinds)
            _correctedLabels[kv.Key] = ColumnKindLabels.ToLabel(kv.Value);
        RebuildTypeGrid();
    }

    private void KindCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0 || sender is not ComboBox cb || cb.DataContext is not ColumnTypeRow row)
            return;
        var label = cb.SelectedItem as string ?? "";
        row.CorrectedLabel = label;
        _correctedLabels[row.ColumnName] = label;
    }

    private void OutputMode_Changed(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded) return;
        RefreshOutputHint();
    }

    private void RefreshOutputHint()
    {
        if (TxtOutputHint == null) return;

        if (string.IsNullOrEmpty(_filePath))
        {
            TxtOutputHint.Text = "（选择文件后显示）";
            return;
        }
        var sourceFile = _filePath!;
        var mode = RbModeWorkbook?.IsChecked == true
            ? ExcelProcessor.OutputModeWorkbook
            : ExcelProcessor.OutputModeFiles;
        if (mode == ExcelProcessor.OutputModeWorkbook)
        {
            var p = ExcelProcessor.MakeTimestampedOutputFile(sourceFile);
            TxtOutputHint.Text = $"将生成：{p}";
        }
        else
        {
            var d = ExcelProcessor.MakeTimestampedOutputDir(sourceFile);
            TxtOutputHint.Text = $"将生成文件夹：{d}";
        }
    }

    private async void RunSplit_Click(object sender, RoutedEventArgs e)
    {
        if (_data == null || string.IsNullOrEmpty(_filePath))
        {
            MessageBox.Show(this, "请先选择并加载 Excel 文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (_splitColumns.Count == 0)
        {
            MessageBox.Show(this, "请至少添加一列用于拆分。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        foreach (var row in _typeRows)
            _correctedLabels[row.ColumnName] = row.CorrectedLabel;

        var sourceFile = _filePath!;
        BtnRun.IsEnabled = false;
        ProgressSplit.Visibility = Visibility.Visible;
        ProgressSplit.IsIndeterminate = true;
        ProgressSplit.Value = 0;
        _setStatus("正在拆分…");
        var progress = new Progress<(int current, int total)>(p =>
        {
            if (p.total <= 0) return;
            ProgressSplit.IsIndeterminate = false;
            ProgressSplit.Value = Math.Min(100, p.current * 100.0 / p.total);
            _setStatus($"正在拆分… ({p.current}/{p.total})");
        });
        try
        {
            var mode = RbModeWorkbook.IsChecked == true
                ? ExcelProcessor.OutputModeWorkbook
                : ExcelProcessor.OutputModeFiles;
            var outputPath = mode == ExcelProcessor.OutputModeWorkbook
                ? ExcelProcessor.MakeTimestampedOutputFile(sourceFile)
                : ExcelProcessor.MakeTimestampedOutputDir(sourceFile);
            var kinds = GetColKindsDict();
            var filtered = ExcelProcessor.ApplyFilters(_data, _filters);
            var result = await System.Threading.Tasks.Task.Run(() =>
                ExcelProcessor.SplitAndExport(
                    filtered, _splitColumns, outputPath, _headerRows, kinds, mode, progress));
            _lastResult = result;
            MenuOpenOutput.IsEnabled = true;
            var msg = mode == ExcelProcessor.OutputModeWorkbook
                ? $"已写入 {result.Count} 个工作表。\n{result.OutputPath}"
                : $"已生成 {result.Count} 个文件。\n{result.OutputPath}";
            ProgressSplit.IsIndeterminate = false;
            ProgressSplit.Value = 100;
            _setStatus($"完成：{result.Count} 项");
            if (MessageBox.Show(this, msg + "\n\n是否打开输出位置？", "拆分完成",
                    MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                OpenPath(result.OutputPath);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, FormatUserError(ex), "拆分失败", MessageBoxButton.OK, MessageBoxImage.Error);
            _setStatus("拆分失败");
        }
        finally
        {
            BtnRun.IsEnabled = true;
            ProgressSplit.Visibility = Visibility.Collapsed;
            ProgressSplit.IsIndeterminate = false;
            ProgressSplit.Value = 0;
        }
    }

    private void OpenOutput_Click(object sender, RoutedEventArgs e)
    {
        if (_lastResult != null) OpenPath(_lastResult.OutputPath);
    }

    private static void OpenPath(string path)
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

    private void About_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(this,
            "Excel 表格拆分工具\n\n" +
            "· 按列拆分（多文件 / 单工作簿）\n" +
            "· 筛选、列类型识别与导出\n\n" +
            "界面：WPF\n" +
            "目标：.NET Framework 4.8（Win7 SP1+）",
            "关于", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void Exit_Click(object sender, RoutedEventArgs e) => Close();

    private void SetFileStatus(bool ok, string text)
    {
        DotFile.Fill = ok ? (Brush)FindResource("SuccessBrush") : Brushes.LightGray;
        LblFileStatus.Text = ok ? text : "未加载文件";
    }

    private void SetSheetStatus(bool ok, string text)
    {
        DotSheet.Fill = ok ? (Brush)FindResource("SuccessBrush") : Brushes.LightGray;
        LblSheetStatus.Text = ok ? text : "无工作表";
    }

    private void _setStatus(string text) => TxtStatus.Text = text;
}
