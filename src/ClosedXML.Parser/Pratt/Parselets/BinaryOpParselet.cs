namespace ClosedXML.Parser.Pratt.Parselets;

internal class BinaryOpParselet<TScalar, T, TContext> : IParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;
    private readonly BinaryOperation _op;
    private readonly int _bp;

    public BinaryOpParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser, BinaryOperation op, int bp)
    {
        _factory = factory;
        _parser = parser;
        _op = op;
        _bp = bp;
    }

    public Node<T> Parse(TContext ctx, Node<T> left, Token op)
    {
        var right = _parser.ParseExpression(ctx, _bp);
        var nodeRange = left.Range
            .ExtendRight(op.Range)
            .ExtendRight(right.Range);

        var node = _factory.BinaryNode(ctx, nodeRange, _op, left, right);
        return new Node<T>(node, nodeRange);
    }

    public int GetBindingPower()
    {
        return _bp;
    }
}

