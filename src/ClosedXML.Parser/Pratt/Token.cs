using System;

namespace ClosedXML.Parser.Pratt;

internal readonly struct Token
{
    public Token(TokenType type, int start, int end)
    {
        Type = type;
        Range = new SymbolRange(start, end);
    }

    public TokenType Type { get; }

    public SymbolRange Range { get; }

    public ReadOnlySpan<char> GetText(string input)
    {
        return input.AsSpan(Range.Start, Range.Length);
    }
}
