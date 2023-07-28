namespace ClosedXML.Parser;

public record Reference3DNode(string FirstSheet, string LastSheet, ReferenceArea Area) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"{FirstSheet}:{LastSheet}!{Area.GetDisplayString(style)}";
    }
}