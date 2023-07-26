namespace ClosedXML.Parser;

public record CellFunctionNode(CellReference Cell) : AstNode
{
    public override string GetDisplayString() => Cell.GetDisplayString();
}