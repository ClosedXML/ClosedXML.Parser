using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// Due to frequency of an area in formulas, the grammar has a token that represents
/// an area in a sheet. This is the DTO from parser to engine. Two corners make an area
/// for A1 notation, but not for R1C1 (has several edge cases).
/// </summary>
public readonly struct ReferenceSymbol
{
    /// <summary>
    /// First reference. First in terms of position in formula, not position
    /// in sheet.
    /// </summary>
    public readonly RowCol First;

    /// <summary>
    /// Second reference. Second in terms of position in formula, not position
    /// in sheet. If area was specified using only one cell, the value is
    /// same as <see cref="First"/>.
    /// </summary>
    public readonly RowCol Second;

    /// <summary>
    /// Create a reference symbol using the two <see cref="RowCol"/> (e.g.
    /// <c>A1:B2</c>) or two columns (e.g. <c>A:D</c>) or two rows (e.g.
    /// <c>7:8</c>).
    /// </summary>
    internal ReferenceSymbol(RowCol first, RowCol second)
    {
        First = first;
        Second = second;
    }

    /// <summary>
    /// Create an area for a single reference.
    /// </summary>
    internal ReferenceSymbol(RowCol rowCol)
    {
        First = rowCol;
        Second = rowCol;
    }

    /// <summary>
    /// Create a new area from a single <see cref="RowCol"/>.
    /// </summary>
    /// <param name="rowType">Row axis type of a reference.</param>
    /// <param name="rowPosition">Row position.</param>
    /// <param name="columnType">Column axis type of a reference.</param>
    /// <param name="columnPosition">Column position.</param>
    internal ReferenceSymbol(ReferenceAxisType rowType, int rowPosition, ReferenceAxisType columnType, int columnPosition)
        : this(new RowCol(rowType, rowPosition, columnType, columnPosition))
    {
    }

    /// <summary>
    /// Create a new area from a single row/column intersection.
    /// </summary>
    /// <param name="rowPosition"><see cref="ReferenceAxisType.Relative"/> row.</param>
    /// <param name="columnPosition"><see cref="ReferenceAxisType.Relative"/> column.</param>
    internal ReferenceSymbol(int rowPosition, int columnPosition)
        : this(new RowCol(ReferenceAxisType.Relative, rowPosition, ReferenceAxisType.Relative, columnPosition))
    {
    }

    /// <summary>
    /// Render area in A1 notation. The content must be a valid content
    /// from A1 token. If both references are same, only one is converted
    /// to display string.
    /// </summary>
    public string GetDisplayStringA1()
    {
        if (First == Second)
            return First.GetDisplayStringA1();

        return new StringBuilder()
            .Append(First.GetDisplayStringA1())
            .Append(':')
            .Append(Second.GetDisplayStringA1())
            .ToString();
    }

    /// <summary>
    /// Render area in R1C1 notation. The content must be a valid content
    /// from R1C1 token. If both references are same, only one is converted
    /// to display string.
    /// </summary>
    public string GetDisplayStringR1C1()
    {
        if (First == Second)
            return First.GetDisplayStringR1C1();

        return new StringBuilder()
            .Append(First.GetDisplayStringR1C1())
            .Append(':')
            .Append(Second.GetDisplayStringR1C1())
            .ToString();
    }
}