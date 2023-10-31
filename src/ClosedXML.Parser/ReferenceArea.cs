using System;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// Due to frequency of an area in formulas, the grammar has a token that represents
/// an area in a sheet. This is the DTO from parser to engine. Two corners make an area
/// for A1 notation, but not for R1C1 (has several edge cases).
/// </summary>
public readonly struct ReferenceArea
{
    /// <summary>
    /// First reference. First in terms of position in formula, not position
    /// in sheet.
    /// </summary>
    public RowCol First { get; }

    /// <summary>
    /// Second reference. Second in terms of position in formula, not position
    /// in sheet. If area was specified using only one cell, the value is
    /// same as <see cref="First"/>.
    /// </summary>
    public RowCol Second { get; }

    /// <summary>
    /// Semantic style of reference.
    /// </summary>
    public ReferenceStyle Style => First.Style;

    /// <summary>
    /// Create a reference symbol using the two <see cref="RowCol"/> (e.g.
    /// <c>A1:B2</c>) or two columns (e.g. <c>A:D</c>) or two rows (e.g.
    /// <c>7:8</c>).
    /// </summary>
    public ReferenceArea(RowCol first, RowCol second)
    {
        if (first.IsA1 ^ second.IsA1)
            throw new ArgumentException("Both RowCol must use same semantic.");

        First = first;
        Second = second;
    }

    /// <summary>
    /// Create an area for a single reference.
    /// </summary>
    internal ReferenceArea(RowCol rowCol)
        : this(rowCol, rowCol)
    {
    }

    /// <summary>
    /// Create a new area from a single <see cref="RowCol"/>.
    /// </summary>
    /// <param name="rowType">Row axis type of a reference.</param>
    /// <param name="rowPosition">Row position.</param>
    /// <param name="columnType">Column axis type of a reference.</param>
    /// <param name="columnPosition">Column position.</param>
    /// <param name="style">Semantic of the reference.</param>
    internal ReferenceArea(ReferenceAxisType rowType, int rowPosition, ReferenceAxisType columnType,
        int columnPosition, ReferenceStyle style)
        : this(new RowCol(rowType, rowPosition, columnType, columnPosition, style))
    {
    }

    /// <summary>
    /// Create a new area from a single row/column intersection.
    /// </summary>
    /// <param name="rowPosition"><see cref="ReferenceAxisType.Relative"/> row.</param>
    /// <param name="columnPosition"><see cref="ReferenceAxisType.Relative"/> column.</param>
    /// <param name="style">Semantic of the reference.</param>
    internal ReferenceArea(int rowPosition, int columnPosition, ReferenceStyle style)
        : this(new RowCol(ReferenceAxisType.Relative, rowPosition, ReferenceAxisType.Relative, columnPosition, style))
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

    /// <summary>
    /// Convert A1 reference to R1C1.
    /// </summary>
    /// <remarks>Assumes reference is in A1.</remarks>
    internal ReferenceArea ToR1C1(int anchorRow, int anchorCol)
    {
        var first = First.ToR1C1(anchorRow, anchorCol);
        if (First != Second)
        {
            var second = Second.ToR1C1(anchorRow, anchorCol);
            return new ReferenceArea(first, second);
        }

        return new ReferenceArea(first, first);
    }

    /// <summary>
    /// Convert R1C1 reference to A1.
    /// </summary>
    /// <remarks>Assumes reference is in R1C1.</remarks>
    internal ReferenceArea ToA1(int anchorRow, int anchorCol)
    {
        var first = First.ToA1(anchorRow, anchorCol);
        if (First != Second)
        {
            var second = Second.ToA1(anchorRow, anchorCol);
            return new ReferenceArea(first, second);
        }

        return new ReferenceArea(first, first);
    }

    internal StringBuilder Append(StringBuilder sb)
    {
        return Style == A1 ? AppendA1(sb) : AppendR1C1(sb);
    }

    /// <summary>
    /// Get reference in A1 notation.
    /// </summary>
    /// <remarks>Assumes reference is in A1.</remarks>
    internal StringBuilder AppendA1(StringBuilder sb)
    {
        First.AppendA1(sb);

        if (First != Second)
        {
            sb.Append(':');
            Second.AppendA1(sb);
        }

        return sb;
    }

    /// <summary>
    /// Get reference in R1C1 notation.
    /// </summary>
    /// <remarks>Assumes reference is in R1C1.</remarks>
    internal StringBuilder AppendR1C1(StringBuilder sb)
    {
        First.AppendR1C1(sb);

        if (First != Second)
        {
            sb.Append(':');
            Second.AppendR1C1(sb);
        }

        return sb;
    }
}