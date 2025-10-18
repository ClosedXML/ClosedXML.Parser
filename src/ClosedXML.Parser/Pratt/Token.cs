namespace ClosedXML.Parser.Pratt;

internal readonly struct Token
{
    public Token(TokenType type, string text)
    {
        Type = type;
        Text = text;
    }

    public TokenType Type { get; }

    public string Text { get; }
}