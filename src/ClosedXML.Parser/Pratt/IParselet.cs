namespace ClosedXML.Parser.Pratt;

internal interface IParselet<T, in TContext>
{
    Node<T> Parse(TContext ctx, Node<T> left, Token op);

    int GetBindingPower();
}
