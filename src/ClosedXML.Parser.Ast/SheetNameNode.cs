namespace ClosedXML.Parser;

public record SheetNameNode(string Sheet, string Name) : AstNode
{
    public override string GetDisplayString()
    {
        return $"[{Sheet}]!{Name}";
    }
}