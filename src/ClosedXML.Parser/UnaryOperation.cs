namespace ClosedXML.Parser;

/// <summary>
/// Unary operations of a formula.
/// </summary>
/// <remarks>Range operations are always after number operations.</remarks>
public enum UnaryOperation
{
    /// <summary>Prefix plus operation.</summary>
    Plus,

    /// <summary>Prefix minus operation.</summary>
    Minus,

    /// <summary>Suffix percent operation.</summary>
    Percent,

    /// <summary>Prefix range intersection operation.</summary>
    Intersect,

    /// <summary>Suffix range spill operation.</summary>
    Spill,
}