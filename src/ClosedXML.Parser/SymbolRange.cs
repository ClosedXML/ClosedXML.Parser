using System;

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

    internal SymbolRange ExtendRight(SymbolRange rangeToRight)
    {
        if (End != rangeToRight.Start)
            throw new InvalidOperationException($"The range end {End} doesn't match start of the range to the right {rangeToRight.Start}.");

        return new SymbolRange(Start, rangeToRight.End);
    }
}
