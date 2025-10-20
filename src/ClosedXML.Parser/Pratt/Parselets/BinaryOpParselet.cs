namespace ClosedXML.Parser.Pratt.Parselets;

internal class BinaryOpParselet<TScalar, TNode, TContext> : IParselet<TNode, TContext>
{
    private readonly IAstFactory<TScalar, TNode, TContext> _factory;
    private readonly Parser<TNode, TContext> _parser;
    private readonly BinaryOperation _op;
    private readonly int _bp;

    public BinaryOpParselet(IAstFactory<TScalar, TNode, TContext> factory, Parser<TNode, TContext> parser, BinaryOperation op, int bp)
    {
        _factory = factory;
        _parser = parser;
        _op = op;
        _bp = bp;
    }

    public TNode Parse(TContext ctx, TNode left, Token op)
    {
        var right = _parser.ParseExpression(ctx, _bp);
        var node = _factory.BinaryNode(ctx, op.Range, _op, left, right); // TODO: Fix binary node range
        return node;
    }

    public int GetBindingPower()
    {
        return _bp;
    }
}

