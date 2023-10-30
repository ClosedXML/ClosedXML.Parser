namespace ClosedXML.Parser;

public record Reference3DNode(string FirstSheet, string LastSheet, ReferenceArea Reference) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        return $"{FirstSheet}:{LastSheet}!{Reference.GetDisplayString(style)}";
    }
}