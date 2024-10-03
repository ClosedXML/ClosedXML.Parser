using System;

namespace ClosedXML.Parser;

/// <summary>
/// A context for modifications of a formula through <see cref="CopyVisitor"/>.
/// </summary>
public class ModContext
{
    /// <summary>
    /// Create a context for modifying formulas. 
    /// </summary>
    [Obsolete("Use overload with sheet parameter.")]
    public ModContext(string formula, int row, int col, bool isA1)
        : this(formula, string.Empty, row, col, isA1)
    {
    }

    /// <summary>
    /// Create a context for modifying formulas. 
    /// </summary>
    public ModContext(string formula, string sheet, int row, int col, bool isA1)
    {
        if (string.IsNullOrWhiteSpace(formula))
            throw new ArgumentException(nameof(formula));

        if (row is < 1 or > RowCol.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(row));

        if (col is < 1 or > RowCol.MaxCol)
            throw new ArgumentOutOfRangeException(nameof(row));

        Formula = formula;
        Sheet = sheet;
        Row = row;
        Col = col;
        IsA1 = isA1;
    }

    /// <summary>
    /// The original formula without any modifications.
    /// </summary>
    public string Formula { get; }

    /// <summary>
    /// Name of the current sheet.
    /// </summary>
    public string Sheet { get; }

    /// <summary>
    /// Absolute row number in a sheet.
    /// </summary>
    public int Row { get; }

    /// <summary>
    /// Absolute column number in a sheet.
    /// </summary>
    public int Col { get; }

    /// <summary>
    /// Should references in formula be A1?
    /// </summary>
    public bool IsA1 { get; }
}
