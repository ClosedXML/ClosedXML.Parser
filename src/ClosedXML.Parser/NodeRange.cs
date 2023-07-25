namespace ClosedXML.Parser;

/// <summary>
/// A range of characters of a symbol in the input. It's not a token, but a range of chars
/// used to create a node. Mostly for A1 - R1C1 formula conversion. 
/// </summary>
public readonly struct NodeRange
{
    /// <summary>
    /// First index in the <c>input</c> that is a part of a node symbol.
    /// </summary>
    public readonly int StartIndex;

    /// <summary>
    /// Length of a symbol. Always at least 1.
    /// </summary>
    public readonly int Length;

    public NodeRange(int startIndex, int length)
    {
        StartIndex = startIndex;
        Length = length;
    }

    internal NodeRange(Token token)
    {
        StartIndex = token.StartIndex;
        Length = token.Length;
    }
}