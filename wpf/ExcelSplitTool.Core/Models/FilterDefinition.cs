using System;

namespace ExcelSplitTool.Core.Models;

public sealed class FilterDefinition
{
    public string Column { get; set; } = "";
    public ColumnKind Kind { get; set; } = ColumnKind.Text;
    public DateTime? Start { get; set; }
    public DateTime? End { get; set; }
    public double? Min { get; set; }
    public double? Max { get; set; }
    public string Text { get; set; } = "";
}
