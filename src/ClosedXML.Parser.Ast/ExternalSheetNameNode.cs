namespace ClosedXML.Parser;

public record ExternalSheetNameNode(int WorkbookIndex, string Sheet, string Name) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"[{WorkbookIndex}]{Sheet}!{Name}";
    }
}