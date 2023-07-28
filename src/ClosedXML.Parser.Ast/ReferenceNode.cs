namespace ClosedXML.Parser;

public record ReferenceNode(ReferenceArea Area) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return Area.GetDisplayString(style);
    }
}