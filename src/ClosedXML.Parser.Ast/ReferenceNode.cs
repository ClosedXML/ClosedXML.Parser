namespace ClosedXML.Parser;

public record ReferenceNode(ReferenceSymbol Reference) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return Reference.GetDisplayString(style);
    }
}