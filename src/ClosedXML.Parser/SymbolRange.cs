namespace ClosedXML.Parser;

/// <summary>
/// A range of a symbol in formula text.
/// </summary>
public readonly struct SymbolRange
{
    /// <summary>
    /// Create a substring of a symbol.
    /// </summary>
    public SymbolRange(int startIndex, int endIndex)
    {
        Start = startIndex;
        End = endIndex;
    }

    /// <summary>
    /// Start index of symbol in formula text.
    /// </summary>
    public int Start { get; }

    /// <summary>
    /// End index of symbol in formula text. Can be outside of text bounds, if symbol ends at the
    /// last char of formula.
    /// </summary>
    public int End { get; }

    /// <summary>
    /// Length of a symbol.
    /// </summary>
    public int Length => End - Start;

    /// <summary>
    /// Get range indexes.
    /// </summary>
    public override string ToString()
    {
        return $"[{Start}:{End}]";
    }
}