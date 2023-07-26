namespace ClosedXML.Parser;

public record LocalReferenceNode(CellArea Reference) : AstNode
{
    public override string GetDisplayString()
    {
        return Reference.GetDisplayString();
    }
}