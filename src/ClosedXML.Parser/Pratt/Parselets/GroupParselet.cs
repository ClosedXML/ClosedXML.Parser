namespace ClosedXML.Parser.Pratt.Parselets;

internal class GroupParselet<TNode, TContext> : IPrefixParselet<TNode, TContext>
{
    private readonly Parser<TNode, TContext> _parser;

    public GroupParselet(Parser<TNode, TContext> parser)
    {
        _parser = parser;
    }

    public TNode Parse(TContext ctx, Token token)
    {
        var node = _parser.ParseExpression(ctx, 0);
        _parser.Consume(TokenType.RightParen);
        return node;
    }
}
