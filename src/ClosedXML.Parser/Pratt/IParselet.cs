namespace ClosedXML.Parser.Pratt;

internal interface IParselet<TNode, in TContext>
{
    TNode Parse(TContext ctx, TNode left, Token op);

    int GetBindingPower();
}
