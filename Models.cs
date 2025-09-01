using System;
using System.Collections.Generic;

namespace Excel_mcp_dotnet;

public class AdvancedFormatting
{
    public FontFormatting? Font { get; set; }
    public FillFormatting? Fill { get; set; }
    public BorderFormatting? Border { get; set; }
    public AlignmentFormatting? Alignment { get; set; }
    public string? NumberFormat { get; set; }
}

public class FontFormatting
{
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string? Color { get; set; }
    public double? Size { get; set; }
    public string? Name { get; set; }
}

public class FillFormatting
{
    public string? BackgroundColor { get; set; }
}

public class BorderFormatting
{
    public string? Color { get; set; }
    public string? Style { get; set; } // "thin", "medium", "thick", etc.
}

public class AlignmentFormatting
{
    public string? Horizontal { get; set; } // "left", "center", "right", etc.
    public string? Vertical { get; set; } // "top", "middle", "bottom", etc.
    public bool? WrapText { get; set; }
}

public class ConditionalFormattingCondition
{
    public string? Operator { get; set; } // "equal", "greaterThan", "lessThan", etc.
    public object? Value { get; set; }
    public string? Formula { get; set; }
}

public class ConditionalFormattingFormat
{
    public string? BackgroundColor { get; set; }
    public string? FontColor { get; set; }
    public bool? Bold { get; set; }
}

public class SortColumn
{
    public int ColumnIndex { get; set; }
    public bool Ascending { get; set; } = true;
}

public class PivotValue
{
    public string? Field { get; set; }
    public string? Function { get; set; } = "sum";
}

public class ValidationCriteria
{
    public List<string>? Values { get; set; }
    public double? MinValue { get; set; }
    public double? MaxValue { get; set; }
    public DateTime? StartDate { get; set; }
    public DateTime? EndDate { get; set; }
    public bool? AllowBlank { get; set; }
    public string? ErrorMessage { get; set; }
    public string? ErrorTitle { get; set; }
    public string? InputMessage { get; set; }
    public string? InputTitle { get; set; }
}