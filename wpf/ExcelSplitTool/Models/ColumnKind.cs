namespace ExcelSplitTool.Models;

public enum ColumnKind
{
    Text,
    Date,
    Number
}

public static class ColumnKindLabels
{
    public static readonly string[] All = { "文本", "日期", "数字" };

    public static string ToLabel(ColumnKind k) => k switch
    {
        ColumnKind.Date => "日期",
        ColumnKind.Number => "数字",
        _ => "文本"
    };

    public static ColumnKind FromLabel(string? label) => label switch
    {
        "日期" => ColumnKind.Date,
        "数字" => ColumnKind.Number,
        _ => ColumnKind.Text
    };
}
