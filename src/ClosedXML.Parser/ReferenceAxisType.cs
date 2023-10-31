namespace ClosedXML.Parser;

/// <summary>
/// The type of content stored in a row or column number of a reference.
/// </summary>
public enum ReferenceAxisType
{
    /// <summary>
    /// Axis is relative. E.g. <c>A5</c> for A1, <c>R[-3]</c> for R1C1.
    /// </summary>
    /// <remarks>Keep 0, so default <c>RowCol</c> is <em>A1</em>.</remarks>
    Relative = 0,

    /// <summary>
    /// Units are absolute. E.g. <c>$A$5</c> for A1, <c>R8C5</c> for R1C1.
    /// </summary>
    Absolute = 1,

    /// <summary>
    /// <para>
    /// The reference axis (row or column) is not specified for reference.
    /// Generally, it means whole axis is used. If the type is <see cref="None"/>,
    /// the value is ignored, but should be 0.
    /// </para>
    /// <para>
    /// Examples:
    /// <list type="bullet">
    ///   <item><c>A:B</c> in A1 doesn't specify row.</item>
    ///   <item><c>R2</c> in R1C1 doesn't specify column.</item>
    /// </list>
    /// </para>
    /// </summary>
    None = 2,
}