using System.Globalization;

namespace ClosedXML.Parser.Pratt.Parselets;

/// <summary>
/// Get a number node from a <see cref="TokenType.Number"/> token.
/// </summary>
/// <remarks>
/// <c>double.Parse</c> parses even <c>NaN</c> or <c>∞</c>, but we can never receive such text
/// from the lexer.
/// </remarks>
internal class NumberParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public NumberParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        _factory = factory;
        _parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
#if NETSTANDARD2_1
        var text = token.GetText(_parser.Input);
#else
        var text = token.GetText(_parser.Input).ToString(); // NetFx has a double whammy, it's slow and gets extra memory to GC
#endif
        var number = double.Parse(text, NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent, CultureInfo.InvariantCulture);
        var node = _factory.NumberNode(ctx, token.Range, number);
        return new Node<T>(node, token.Range);
    }
}
