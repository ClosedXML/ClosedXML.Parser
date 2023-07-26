namespace ClosedXML.Parser;

public record ExternalFunctionNode(int WorkbookIndex, string? Sheet, string Name) : AstNode
{
    public override string GetDisplayString()
    {
        return $"[{WorkbookIndex}]{Sheet}!{Name}";
    }
}