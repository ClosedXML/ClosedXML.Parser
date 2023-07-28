namespace ClosedXML.Parser;

public record ExternalNameNode(int WorkbookIndex, string Name) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"[{WorkbookIndex}]!{Name}";
    }
}