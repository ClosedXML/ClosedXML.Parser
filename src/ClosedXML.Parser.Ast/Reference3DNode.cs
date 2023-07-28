namespace ClosedXML.Parser;

public record Reference3DNode(string FirstSheet, string LastSheet, ReferenceArea Area) : AstNode
{
    public override string GetDisplayString()
    {
        return $"{FirstSheet}:{LastSheet}!{Area.GetDisplayString()}";
    }
}