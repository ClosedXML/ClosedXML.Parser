namespace ClosedXML.Parser;

public record ExternalReferenceNode(int WorkbookIndex, CellArea Reference) : AstNode
{
    public override string GetDisplayString()
    {
        return $"[{WorkbookIndex}]{Reference}!{Reference.GetDisplayString()}";
    }
};