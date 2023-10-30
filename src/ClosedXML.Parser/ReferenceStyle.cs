namespace ClosedXML.Parser;

/// <summary>
/// Style of referencing areas in a worksheet.
/// </summary>
public enum ReferenceStyle
{
    /// <summary>
    /// The reference (<see cref="ReferenceSymbol"/> or <see cref="RowCol"/>)
    /// uses <em>A1</em> semantic. Even relative references start from
    /// <c>[1,1]</c>, but relative references move when cells move.
    /// </summary>
    A1,

    /// <summary>
    /// The reference (<see cref="ReferenceSymbol"/> or <see cref="RowCol"/>)
    /// uses <em>R1C1</em> semantic. Relative references are relative to
    /// the cell that contains the reference, not <c>[1,1]</c>.
    /// </summary>
    R1C1,
}