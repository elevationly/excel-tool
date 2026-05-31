using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelDataReader;
using ExcelSplitTool.Core.Models;

namespace ExcelSplitTool.Core.Services;

public static class ExcelProcessor
{
    public const string OutputFolderLabel = "拆分表格";
    public const string OutputModeFiles = "files";
    public const string OutputModeWorkbook = "workbook";

    private static readonly Regex InvalidFilenameChars = new(@"[<>:""/\\|?*\x00-\x1f]");
    private static readonly Regex InvalidSheetChars = new(@"[\\/*?:\[\]]");
    private const int MaxSheetNameLen = 31;

    private static readonly string[] DateKeywords = { "日期", "时间", "年月" };
    private static readonly string[] NumberKeywords = { "投资", "金额", "万", "数量", "容量", "长度", "含税", "不含税" };
    private static readonly string[] TextIdKeywords = { "序号", "编号", "文号", "凭证", "WBS", "物料", "人", "名称", "描述" };
    private static readonly string[] NonDateKeywords = { "序号", "编号", "年份", "批次", "性质", "类型", "分类" };
    private static readonly string[] SubheaderHints = { "计划", "实际", "开工", "竣工", "送审", "审定", "关闭", "决算", "容量", "长度" };

    private static bool _codePagesRegistered;
    private static string? _codePagesError;

    /// <summary>延迟注册编码提供程序，避免静态构造函数在部分 Win7 环境下失败。</summary>
    private static void EnsureCodePages()
    {
        if (_codePagesRegistered) return;
        _codePagesRegistered = true;
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        catch (Exception ex)
        {
            _codePagesError = ex.Message;
        }
    }

    public static List<string> ListSheetNames(string path)
    {
        EnsureCodePages();
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var names = new List<string>();
        do { names.Add(reader.Name); } while (reader.NextResult());
        return names;
    }

    public static ExcelSheet ReadExcelWithHeaders(string path, string sheetName)
    {
        var raw = ReadRawGrid(path, sheetName);
        if (raw.Count == 0) throw new InvalidOperationException("表格为空");

        var headerCount = DetectHeaderRowCount(raw);
        var headerRows = raw.Take(headerCount).Select(r => r.ToList()).ToList();
        var dataRows = raw.Skip(headerCount).ToList();
        var columns = BuildColumnLabels(headerRows);
        var table = new DataTable();
        var colCount = dataRows.Count > 0 ? dataRows.Max(r => r.Count) : columns.Count;
        for (var i = 0; i < colCount; i++)
        {
            var name = i < columns.Count ? columns[i] : $"列{i + 1}";
            table.Columns.Add(name, typeof(object));
        }
        foreach (var row in dataRows)
        {
            var dr = table.NewRow();
            for (var i = 0; i < colCount; i++)
                dr[i] = i < row.Count ? row[i] ?? DBNull.Value : DBNull.Value;
            table.Rows.Add(dr);
        }
        return new ExcelSheet(table, headerRows);
    }

    private static List<List<object?>> ReadRawGrid(string path, string sheetName)
    {
        EnsureCodePages();
        if (_codePagesError != null && path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "读取 .xls 需要编码组件，当前环境加载失败。请安装运行库后重试，或改用 .xlsx 文件。\n" + _codePagesError);

        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        if (!string.IsNullOrEmpty(sheetName))
        {
            while (reader.Name != sheetName && reader.NextResult()) { }
        }
        var rows = new List<List<object?>>();
        while (reader.Read())
        {
            var row = new List<object?>();
            for (var i = 0; i < reader.FieldCount; i++)
                row.Add(reader.IsDBNull(i) ? null : reader.GetValue(i));
            rows.Add(row);
        }
        return rows;
    }

    private static string CellStr(object? value)
    {
        if (value == null || value == DBNull.Value) return "";
        return Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? "";
    }

    private static bool LooksLikeDataValue(object? value)
    {
        var s = CellStr(value);
        if (string.IsNullOrEmpty(s)) return false;
        if (s.Length >= 15) return true;
        var digits = s.Replace(".", "").Replace("-", "");
        return digits.All(char.IsDigit) && digits.Length >= 8;
    }

    private static bool LooksLikeSubheader(object? value)
    {
        var s = CellStr(value);
        if (string.IsNullOrEmpty(s) || s.Length > 12) return false;
        return SubheaderHints.Any(h => s.Contains(h));
    }

    private static int DetectHeaderRowCount(List<List<object?>> raw)
    {
        if (raw.Count < 2) return 1;
        var row0 = raw[0];
        var row1 = raw[1];
        var r1NonNull = row1.Count(v => v != null && v != DBNull.Value);
        if (r1NonNull == 0) return 1;

        var sampleN = Math.Min(12, row0.Count);
        var dataLike = 0;
        for (var i = 0; i < sampleN; i++)
        {
            if (i >= row0.Count || i >= row1.Count) continue;
            if (row0[i] != null && row1[i] != null && LooksLikeDataValue(row1[i]))
                dataLike++;
        }
        if (dataLike >= 3) return 1;

        var r1Texts = row1.Where(v => !string.IsNullOrEmpty(CellStr(v))).Select(CellStr).ToList();
        var subheaderLike = row1.Count(LooksLikeSubheader);
        var r0NonNull = Math.Max(row0.Count(v => v != null && v != DBNull.Value), 1);

        if (r1NonNull < r0NonNull * 0.85 && subheaderLike >= 3) return 2;
        if (subheaderLike >= Math.Max(3, (int)(r1Texts.Count * 0.25)) && dataLike == 0) return 2;
        return 1;
    }

    private static List<string> DedupeNames(List<string> names)
    {
        var used = new Dictionary<string, int>();
        var result = new List<string>();
        foreach (var name in names)
        {
            var baseName = string.IsNullOrEmpty(name) ? "未命名" : name;
            if (used.TryGetValue(baseName, out var n))
            {
                used[baseName] = n + 1;
                result.Add($"{baseName}_{n + 1}");
            }
            else
            {
                used[baseName] = 0;
                result.Add(baseName);
            }
        }
        return result;
    }

    public static List<string> BuildColumnLabels(List<List<object?>> headerRows)
    {
        if (headerRows.Count == 1)
        {
            var names = headerRows[0].Select((v, i) =>
            {
                var s = CellStr(v);
                return string.IsNullOrEmpty(s) ? $"列{i + 1}" : s;
            }).ToList();
            return DedupeNames(names);
        }

        var row0 = headerRows[0];
        var row1 = headerRows[1];
        var ncol = Math.Max(row0.Count, row1.Count);
        var names2 = new List<string>();
        string? lastParent = null;
        for (var i = 0; i < ncol; i++)
        {
            var h0 = i < row0.Count ? row0[i] : null;
            var h1 = i < row1.Count ? row1[i] : null;
            var s0 = CellStr(h0);
            if (!string.IsNullOrEmpty(s0)) lastParent = s0;
            var s1 = CellStr(h1);
            string name;
            if (!string.IsNullOrEmpty(s1)) name = s1;
            else if (!string.IsNullOrEmpty(s0)) name = s0;
            else if (!string.IsNullOrEmpty(lastParent)) name = lastParent;
            else name = $"列{i + 1}";
            names2.Add(name);
        }
        return DedupeNames(names2);
    }

    public static Dictionary<string, ColumnKind> DetectAllColumnKinds(DataTable df) =>
        df.Columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => InferColumnKind(df, c.ColumnName));

    public static ColumnKind InferColumnKind(DataTable df, string colName)
    {
        List<object?> series = df.Columns.Contains(colName)
            ? df.AsEnumerable().Select(r => (object?)r[colName]).ToList()
            : new List<object?>();

        if (colName.EndsWith("人") || (colName.Contains("制单") && !colName.Contains("日期") && !colName.Contains("时间")))
            return ColumnKind.Text;
        if (TextIdKeywords.Any(k => colName.Contains(k))) return ColumnKind.Text;
        if (NonDateKeywords.Any(k => colName.Contains(k)))
        {
            if (colName.Contains("日期") || (colName.Contains("时间") && !colName.Contains("年份"))) { }
            else if (colName.Contains("年份")) return ColumnKind.Number;
            else return ColumnKind.Text;
        }
        if (colName.Contains("日期")) return ColumnKind.Date;
        if (colName.Contains("时间") && !colName.Contains("年份")) return ColumnKind.Date;
        if (NumberKeywords.Any(k => colName.Contains(k))) return ColumnKind.Number;
        if (SeriesLooksLikeRealDates(series)) return ColumnKind.Date;

        var sampleList = series.Where(v => v != null && v != DBNull.Value).Take(200).ToList();
        if (sampleList.Count > 0)
        {
            var numericCount = sampleList.Count(v => ToDouble(v).HasValue);
            if (numericCount >= sampleList.Count * 0.6) return ColumnKind.Number;
        }
        return ColumnKind.Text;
    }

    private static bool SeriesLooksLikeRealDates(List<object?> sample)
    {
        var items = sample.Where(v => v != null && v != DBNull.Value).Take(200).ToList();
        if (items.Count == 0) return false;

        if (items[0] is DateTime)
        {
            var years = items.OfType<DateTime>().Select(d => d.Year).ToList();
            return years.Count(y => y >= 1990) >= years.Count * 0.6;
        }

        var asStr = items.Select(v => CellStr(v)).Where(s => s.Length > 0).ToList();
        if (asStr.Count == 0) return false;
        if (asStr.Count(s => Regex.IsMatch(s, @"^\d{5,12}$")) >= asStr.Count * 0.6) return false;
        if (asStr.Count(s => Regex.IsMatch(s, @"^0\d+$")) >= asStr.Count * 0.3) return false;
        if (asStr.Count(s => Regex.IsMatch(s, @"^\d{4}[/\-年]\d{1,2}[/\-月]\d{1,2}")) >= asStr.Count * 0.6) return true;
        if (asStr.Count(s => Regex.IsMatch(s, @"^\d{4}[/\-]\d{1,2}[/\-]\d{1,2}")) >= asStr.Count * 0.6) return true;

        var numeric = items.Select(ToDouble).Where(d => d.HasValue).Select(d => d!.Value).ToList();
        if (numeric.Count >= items.Count * 0.6)
            return numeric.Count(n => n > 10000) >= numeric.Count * 0.6;

        var parsed = items.Select(ToDateTime).Where(d => d.HasValue).Select(d => d!.Value).ToList();
        if (parsed.Count < items.Count * 0.6) return false;
        if (parsed.Count(y => y.Year == 1970) > parsed.Count * 0.3) return false;
        return parsed.Count(y => y.Year >= 1990) >= parsed.Count * 0.6;
    }

    private static double? ToDouble(object? v)
    {
        if (v == null || v == DBNull.Value) return null;
        if (v is double d) return d;
        if (v is float f) return f;
        if (v is int or long or decimal) return Convert.ToDouble(v);
        return double.TryParse(CellStr(v), NumberStyles.Any, CultureInfo.InvariantCulture, out var r) ? r : null;
    }

    private static DateTime? ToDateTime(object? v)
    {
        if (v == null || v == DBNull.Value) return null;
        if (v is DateTime dt) return dt;
        if (double.TryParse(CellStr(v), out var serial) && serial > 10000)
        {
            try { return DateTime.FromOADate(serial); } catch { }
        }
        if (DateTime.TryParse(CellStr(v), CultureInfo.CurrentCulture, DateTimeStyles.None, out var parsed))
            return parsed;
        return null;
    }

    public static DataTable ApplyFilters(DataTable df, IList<FilterDefinition> filters)
    {
        var rows = df.AsEnumerable().ToList();
        foreach (var f in filters)
        {
            if (!df.Columns.Contains(f.Column)) continue;
            rows = rows.Where(r => MatchFilter(r, f)).ToList();
        }
        return rows.Count == 0 ? df.Clone() : rows.CopyToDataTable();
    }

    private static bool MatchFilter(DataRow r, FilterDefinition f)
    {
        var val = r[f.Column];
        switch (f.Kind)
        {
            case ColumnKind.Date:
                if (!f.Start.HasValue && !f.End.HasValue) return true;
                var d = ToDateTime(val);
                if (!d.HasValue) return false;
                var day = d.Value.Date;
                if (f.Start.HasValue && day < f.Start.Value.Date) return false;
                if (f.End.HasValue && day > f.End.Value.Date) return false;
                return true;
            case ColumnKind.Number:
                var n = ToDouble(val);
                if (!n.HasValue) return false;
                if (f.Min.HasValue && n < f.Min) return false;
                if (f.Max.HasValue && n > f.Max) return false;
                return true;
            default:
                var text = (f.Text ?? "").Trim();
                if (string.IsNullOrEmpty(text)) return true;
                return CellStr(val).Contains(text);
        }
    }

    public static string SanitizeFilename(string name, int maxLen = 80)
    {
        var s = InvalidFilenameChars.Replace(name.Trim(), "_").Trim(' ', '.');
        if (string.IsNullOrEmpty(s)) s = "未命名";
        if (s.Length > maxLen) s = s.Substring(0, maxLen);
        return s;
    }

    public static string MakeTimestampedOutputDir(string sourceFile, string? baseDir = null)
    {
        var src = new FileInfo(sourceFile);
        var parent = string.IsNullOrEmpty(baseDir) ? src.DirectoryName! : baseDir;
        var stem = SanitizeFilename(Path.GetFileNameWithoutExtension(src.Name), 60);
        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var outDir = Path.Combine(parent, $"{stem}_{OutputFolderLabel}_{ts}");
        Directory.CreateDirectory(outDir);
        return outDir;
    }

    public static string MakeTimestampedOutputFile(string sourceFile, string? baseDir = null)
    {
        var src = new FileInfo(sourceFile);
        var parent = string.IsNullOrEmpty(baseDir) ? src.DirectoryName! : baseDir;
        var stem = SanitizeFilename(Path.GetFileNameWithoutExtension(src.Name), 50);
        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(parent, $"{stem}_{OutputFolderLabel}_{ts}.xlsx");
    }

    public static SplitExportResult SplitAndExport(
        DataTable df,
        IList<string> splitColumns,
        string outputPath,
        List<List<object?>>? headerRows,
        Dictionary<string, ColumnKind>? colKinds,
        string mode,
        IProgress<(int current, int total)>? progress = null)
    {
        return mode == OutputModeWorkbook
            ? SplitToWorkbook(df, splitColumns, outputPath, headerRows, colKinds, progress)
            : SplitToFiles(df, splitColumns, outputPath, headerRows, colKinds, progress);
    }

    private static SplitExportResult SplitToFiles(
        DataTable df, IList<string> splitColumns, string outputDir,
        List<List<object?>>? headerRows, Dictionary<string, ColumnKind>? colKinds,
        IProgress<(int current, int total)>? progress)
    {
        Directory.CreateDirectory(outputDir);
        var groups = GroupFrames(df, splitColumns).ToList();
        if (groups.Count == 0) throw new InvalidOperationException("没有可写入的数据");
        var count = 0;
        foreach (var (key, group) in groups)
        {
            var parts = key.Select(k => SanitizeFilename(KeyLabel(k))).ToArray();
            var fileName = string.Join("_", parts) + ".xlsx";
            ExportDataTable(group, Path.Combine(outputDir, fileName), headerRows, colKinds);
            count++;
            progress?.Report((count, groups.Count));
        }
        return new SplitExportResult { Count = count, OutputPath = outputDir, Mode = OutputModeFiles };
    }

    private static SplitExportResult SplitToWorkbook(
        DataTable df, IList<string> splitColumns, string outputFile,
        List<List<object?>>? headerRows, Dictionary<string, ColumnKind>? colKinds,
        IProgress<(int current, int total)>? progress)
    {
        var groups = GroupFrames(df, splitColumns).ToList();
        if (groups.Count == 0) throw new InvalidOperationException("没有可写入的数据");
        using var wb = new XLWorkbook();
        var used = new HashSet<string>();
        var count = 0;
        foreach (var (key, group) in groups)
        {
            var parts = key.Select(k => SanitizeFilename(KeyLabel(k), 20)).ToArray();
            var title = SanitizeSheetName(string.Join("_", parts), used);
            var (exportDf, kinds) = PrepareForExport(group, colKinds);
            var ws = wb.Worksheets.Add(title);
            PopulateWorksheet(ws, exportDf, kinds, headerRows);
            count++;
            progress?.Report((count, groups.Count));
        }
        SaveWorkbookSafe(wb, outputFile);
        return new SplitExportResult { Count = count, OutputPath = outputFile, Mode = OutputModeWorkbook };
    }

    private static IEnumerable<(object?[] Key, DataTable Frame)> GroupFrames(DataTable df, IList<string> splitColumns)
    {
        if (splitColumns == null || splitColumns.Count == 0)
            throw new ArgumentException("请至少添加一列用于拆分");
        var missing = splitColumns.Where(c => !df.Columns.Contains(c)).ToList();
        if (missing.Count > 0) throw new KeyNotFoundException($"列不存在: {string.Join(", ", missing)}");

        var groups = df.AsEnumerable().GroupBy(r =>
            string.Join("\u001f", splitColumns.Select(c => KeyLabel(r[c]))));
        foreach (var g in groups)
        {
            var first = g.First();
            var key = splitColumns.Select(c => first[c]).ToArray();
            yield return (key, g.CopyToDataTable());
        }
    }

    private static string KeyLabel(object? k) =>
        k == null || k == DBNull.Value ? "空值" : CellStr(k);

    private static string SanitizeSheetName(string name, HashSet<string> used)
    {
        var s = InvalidSheetChars.Replace(name.Trim(), "_").Trim(' ', '.');
        if (string.IsNullOrEmpty(s)) s = "Sheet";
        if (s.Length > MaxSheetNameLen) s = s.Substring(0, MaxSheetNameLen);
        var baseName = s;
        var n = 1;
        while (used.Contains(s))
        {
            var suffix = $"_{n}";
            s = baseName.Length + suffix.Length > MaxSheetNameLen
                ? baseName.Substring(0, MaxSheetNameLen - suffix.Length) + suffix
                : baseName + suffix;
            n++;
        }
        used.Add(s);
        return s;
    }

    private static void ExportDataTable(
        DataTable df, string path,
        List<List<object?>>? headerRows,
        Dictionary<string, ColumnKind>? colKinds)
    {
        using var wb = new XLWorkbook();
        var (exportDf, kinds) = PrepareForExport(df, colKinds);
        var ws = wb.Worksheets.Add("Sheet1");
        PopulateWorksheet(ws, exportDf, kinds, headerRows);
        SaveWorkbookSafe(wb, path);
    }

    private static (DataTable Export, Dictionary<string, ColumnKind> Kinds) PrepareForExport(
        DataTable df, Dictionary<string, ColumnKind>? colKinds)
    {
        var kinds = new Dictionary<string, ColumnKind>();
        var export = df.Clone();
        foreach (DataColumn col in df.Columns)
        {
            var name = col.ColumnName;
            var kind = colKinds != null && colKinds.TryGetValue(name, out var k)
                ? k
                : InferColumnKind(df, name);
            kinds[name] = kind;
        }
        foreach (DataRow src in df.Rows)
        {
            var dr = export.NewRow();
            foreach (DataColumn col in df.Columns)
            {
                var name = col.ColumnName;
                var v = src[col];
                dr[name] = FormatExportValue(v, kinds[name], name);
            }
            export.Rows.Add(dr);
        }
        return (export, kinds);
    }

    private static object? FormatExportValue(object? value, ColumnKind kind, string colName)
    {
        if (value == null || value == DBNull.Value) return null;
        switch (kind)
        {
            case ColumnKind.Date:
                var d = ToDateTime(value);
                return d.HasValue ? d.Value.ToString("yyyy/MM/dd") : null;
            case ColumnKind.Number:
                var n = ToDouble(value);
                if (!n.HasValue) return null;
                if (colName.Contains("年份") && Math.Abs(n.Value - Math.Round(n.Value)) < 1e-9)
                    return (int)n.Value;
                return n.Value;
            default:
                return ToTextExportValue(value);
        }
    }

    private static string? ToTextExportValue(object? value)
    {
        if (value == null || value == DBNull.Value) return null;
        if (value is double d && Math.Abs(d - Math.Round(d)) < 1e-9)
            return ((long)Math.Round(d)).ToString();
        var s = CellStr(value);
        if (s.EndsWith(".0", StringComparison.Ordinal) && s.Length > 2)
        {
            var core = s.Substring(0, s.Length - 2);
            if (core.All(char.IsDigit)) return core;
        }
        return s;
    }

    private static void PopulateWorksheet(
        IXLWorksheet ws, DataTable exportDf,
        Dictionary<string, ColumnKind> colKinds,
        List<List<object?>>? headerRows)
    {
        var ncol = exportDf.Columns.Count;
        var columns = exportDf.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
        int startRow;
        if (headerRows != null && headerRows.Count > 0)
            startRow = WriteMergedHeaders(ws, headerRows, ncol);
        else
        {
            for (var ci = 0; ci < ncol; ci++)
                ws.Cell(1, ci + 1).Value = columns[ci];
            startRow = 2;
        }
        var ri = startRow;
        foreach (DataRow row in exportDf.Rows)
        {
            for (var ci = 0; ci < ncol; ci++)
                WriteDataCell(ws, ri, ci + 1, row[ci], colKinds[columns[ci]]);
            ri++;
        }
        ApplyWorksheetStyle(ws);
    }

    private static int WriteMergedHeaders(IXLWorksheet ws, List<List<object?>> headerRows, int ncol)
    {
        if (headerRows.Count == 1)
        {
            for (var ci = 0; ci < ncol; ci++)
            {
                var val = ci < headerRows[0].Count ? headerRows[0][ci] : null;
                ws.Cell(1, ci + 1).Value = val?.ToString() ?? "";
            }
            return 2;
        }

        var row0 = headerRows[0].Concat(Enumerable.Repeat<object?>(null, Math.Max(0, ncol - headerRows[0].Count))).Take(ncol).ToList();
        var row1 = headerRows[1].Concat(Enumerable.Repeat<object?>(null, Math.Max(0, ncol - headerRows[1].Count))).Take(ncol).ToList();
        var row0Ffill = new List<string?>();
        string? last = null;
        for (var i = 0; i < ncol; i++)
        {
            var s = CellStr(row0[i]);
            if (!string.IsNullOrEmpty(s)) last = s;
            row0Ffill.Add(last);
        }

        for (var ci = 0; ci < ncol; ci++)
        {
            if (!string.IsNullOrEmpty(CellStr(row1[ci]))) continue;
            var val = CellStr(row0[ci]);
            if (string.IsNullOrEmpty(val)) val = row0Ffill[ci] ?? "";
            ws.Cell(1, ci + 1).Value = val;
            ws.Range(1, ci + 1, 2, ci + 1).Merge();
        }

        var c = 0;
        while (c < ncol)
        {
            if (string.IsNullOrEmpty(CellStr(row1[c]))) { c++; continue; }
            var parent = row0Ffill[c] ?? "";
            var startCi = c;
            while (c < ncol && !string.IsNullOrEmpty(CellStr(row1[c])) && (row0Ffill[c] ?? "") == parent)
            {
                ws.Cell(2, c + 1).Value = CellStr(row1[c]);
                c++;
            }
            var endCi = c - 1;
            if (string.IsNullOrEmpty(parent)) continue;
            ws.Cell(1, startCi + 1).Value = parent;
            if (endCi > startCi)
                ws.Range(1, startCi + 1, 1, endCi + 1).Merge();
        }
        return 3;
    }

    private static void WriteDataCell(IXLWorksheet ws, int row, int col, object? value, ColumnKind kind)
    {
        var cell = ws.Cell(row, col);
        if (value == null || value == DBNull.Value) { cell.Value = ""; return; }
        if (kind == ColumnKind.Text)
        {
            cell.Style.NumberFormat.Format = "@";
            cell.Value = ToTextExportValue(value) ?? "";
        }
        else if (kind == ColumnKind.Number && value is double or int or long or decimal)
            cell.Value = Convert.ToDouble(value);
        else
            cell.Value = value.ToString();
    }

    private static void ApplyWorksheetStyle(IXLWorksheet ws)
    {
        var used = ws.RangeUsed();
        if (used == null) return;
        foreach (var row in used.Rows())
        {
            ws.Row(row.RowNumber()).Height = 30;
            foreach (var cell in row.Cells())
            {
                cell.Style.Font.FontName = "Arial";
                cell.Style.Font.FontSize = 10;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Alignment.WrapText = true;
            }
        }
        var colMax = new Dictionary<int, double>();
        foreach (var cell in used.Cells())
        {
            if (cell.IsEmpty()) continue;
            var text = cell.GetString();
            if (string.IsNullOrEmpty(text)) continue;
            var w = DisplayWidth(text);
            var ci = cell.Address.ColumnNumber;
            if (!colMax.ContainsKey(ci) || colMax[ci] < w) colMax[ci] = w;
        }
        foreach (var kv in colMax)
            ws.Column(kv.Key).Width = Math.Min(Math.Max(kv.Value + 2, 8), 80);
    }

    private static double DisplayWidth(string text)
    {
        double w = 0;
        foreach (var ch in text)
            w += ch > 127 ? 2.0 : 1.0;
        return w;
    }

    private static void SaveWorkbookSafe(XLWorkbook wb, string path)
    {
        var dir = Path.GetDirectoryName(path)!;
        Directory.CreateDirectory(dir);
        var tmp = Path.Combine(dir, Path.GetRandomFileName() + ".xlsx");
        try
        {
            wb.SaveAs(tmp);
            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
        }
        finally
        {
            if (File.Exists(tmp)) try { File.Delete(tmp); } catch { /* ignore */ }
        }
    }
}
