namespace ClosedXML.Parser.Pratt.Parselets;

internal class GroupParselet<T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly Parser<T, TContext> _parser;

    public GroupParselet(Parser<T, TContext> parser)
    {
        _parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token leftParen)
    {
        var node = _parser.ParseExpression(ctx, 0);
        var rightParen = _parser.Consume(TokenType.RightParen);
        return node.ExtendLeft(leftParen).ExtendRight(rightParen);
    }
}
