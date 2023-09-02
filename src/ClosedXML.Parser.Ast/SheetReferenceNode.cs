namespace ClosedXML.Parser;

public record SheetReferenceNode(string Sheet, ReferenceSymbol Reference) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"{Sheet}!{Reference.GetDisplayString(style)}";
    }
}