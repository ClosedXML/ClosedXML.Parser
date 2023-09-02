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
    /// First reference. First in terms of position in formula, not position in sheet.
    /// </summary>
    public readonly RowCol First;

    /// <summary>
    /// Second reference. Second in terms of position in formula, not position in sheet.
    /// If area was specified using only one corner, the value is same as <see cref="First"/>.
    /// </summary>
    public readonly RowCol Second;

    public ReferenceSymbol(RowCol first, RowCol second)
    {
        First = first;
        Second = second;
    }

    /// <summary>
    /// Create an area for a single reference.
    /// </summary>
    public ReferenceSymbol(RowCol rowCol)
    {
        First = rowCol;
        Second = rowCol;
    }

    /// <summary>
    /// Create a new area from a single <see cref="RowCol"/>.
    /// </summary>
    /// <param name="columnType">Column axis type of a reference.</param>
    /// <param name="columnPosition">Column position.</param>
    /// <param name="rowType">Row axis type of a reference.</param>
    /// <param name="rowPosition">Row position.</param>
    public ReferenceSymbol(ReferenceAxisType columnType, int columnPosition, ReferenceAxisType rowType, int rowPosition)
        : this(new RowCol(columnType, columnPosition, rowType, rowPosition))
    {
    }

    /// <summary>
    /// Create a new area from a single row/column intersection.
    /// </summary>
    /// <param name="columnPosition"><see cref="ReferenceAxisType.Relative"/> column.</param>
    /// <param name="rowPosition"><see cref="ReferenceAxisType.Relative"/> row.</param>
    public ReferenceSymbol(int columnPosition, int rowPosition)
        : this(new RowCol(ReferenceAxisType.Relative, columnPosition, ReferenceAxisType.Relative, rowPosition))
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