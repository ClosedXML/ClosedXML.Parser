using System;

namespace ClosedXML.Parser.Pratt;

internal readonly struct Token
{
    private readonly SymbolRange _text;

    public Token(TokenType type, int start, int end)
    {
        Type = type;
        _text = new SymbolRange(start, end);
    }

    public TokenType Type { get; }

    public ReadOnlySpan<char> GetText(string input)
    {
        return input.AsSpan(_text.Start, _text.Length);
    }
}