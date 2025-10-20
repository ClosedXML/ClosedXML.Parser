namespace ClosedXML.Parser.Pratt;

internal interface IPrefixParselet<out TNode, in TContext>
{
    TNode Parse(TContext ctx, Token token);
}
