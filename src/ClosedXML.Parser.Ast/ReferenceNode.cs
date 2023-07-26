namespace ClosedXML.Parser;

public record ReferenceNode(CellArea Reference) : AstNode
{
    public override string GetDisplayString()
    {
        return Reference.GetDisplayString();
    }
}