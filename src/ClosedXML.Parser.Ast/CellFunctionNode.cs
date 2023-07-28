namespace ClosedXML.Parser;

public record CellFunctionNode(Reference Reference) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style) => Reference.GetDisplayString(style);
}