namespace ClosedXML.Parser;

public record SheetReferenceNode(string Sheet, ReferenceArea Area) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"{Sheet}!{Area.GetDisplayString(style)}";
    }
}