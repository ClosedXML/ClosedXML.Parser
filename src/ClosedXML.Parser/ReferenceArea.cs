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
    /// First reference. First in terms of position in formula, not position in sheet.
    /// </summary>
    public readonly Reference First;

    /// <summary>
    /// Second reference. Second in terms of position in formula, not position in sheet.
    /// </summary>
    public readonly Reference Second;

    public ReferenceArea(Reference first, Reference second)
    {
        First = first;
        Second = second;
    }

    /// <summary>
    /// Create an area for a single reference.
    /// </summary>
    public ReferenceArea(Reference reference)
    {
        First = reference;
        Second = reference;
    }

    /// <summary>
    /// Create a new area from a single <see cref="Reference"/>.
    /// </summary>
    /// <param name="columnType">Column axis type of a reference.</param>
    /// <param name="columnPosition">Column position.</param>
    /// <param name="rowType">Row axis type of a reference.</param>
    /// <param name="rowPosition">Row position.</param>
    public ReferenceArea(ReferenceAxisType columnType, int columnPosition, ReferenceAxisType rowType, int rowPosition)
        : this(new Reference(columnType, columnPosition, rowType, rowPosition))
    {
    }

    /// <summary>
    /// Create a new area from a single row/column intersection.
    /// </summary>
    /// <param name="columnPosition"><see cref="ReferenceAxisType.Relative"/> column.</param>
    /// <param name="rowPosition"><see cref="ReferenceAxisType.Relative"/> row.</param>
    public ReferenceArea(int columnPosition, int rowPosition)
        : this(new Reference(ReferenceAxisType.Relative, columnPosition, ReferenceAxisType.Relative, rowPosition))
    {
    }

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
}