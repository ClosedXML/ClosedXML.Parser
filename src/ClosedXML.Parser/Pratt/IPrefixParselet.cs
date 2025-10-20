namespace ClosedXML.Parser.Pratt;

internal interface IPrefixParselet<T, in TContext>
{
    Node<T> Parse(TContext ctx, Token token);
}
