namespace ClosedXML.Parser;

public record ReferenceNode(ReferenceArea Area) : AstNode
{
    public override string GetDisplayString()
    {
        return Area.GetDisplayString();
    }
}