using System.Collections.Generic;
using System.Data;

namespace ExcelSplitTool.Models;

public sealed class ExcelSheet
{
    public DataTable Data { get; }
    public List<List<object?>> HeaderRows { get; }

    public ExcelSheet(DataTable data, List<List<object?>> headerRows)
    {
        Data = data;
        HeaderRows = headerRows;
    }
}

public sealed class SplitExportResult
{
    public int Count { get; set; }
    public string OutputPath { get; set; } = "";
    public string Mode { get; set; } = "files";
}
