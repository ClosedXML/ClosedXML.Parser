namespace ClosedXML.Parser;

public record SheetReferenceNode(string Sheet, ReferenceArea Area) : AstNode
{
    public override string GetDisplayString()
    {
        return $"{Sheet}!{Area.GetDisplayString()}";
    }
}