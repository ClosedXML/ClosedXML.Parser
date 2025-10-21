namespace ClosedXML.Parser;

public record SheetNameNode(string Sheet, string Name) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        var sheet = NameUtils.ShouldQuote(Sheet) ? '\'' + Sheet.Replace("'", "''") + '\'' : Sheet;
        return $"{sheet}!{Name}";
    }
}
