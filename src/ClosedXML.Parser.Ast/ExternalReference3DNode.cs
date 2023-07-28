namespace ClosedXML.Parser;

public record ExternalReference3DNode(int WorkbookIndex, string FirstSheet, string LastSheet, ReferenceArea Area) : AstNode
{
    public override string GetDisplayString()
    {
        return $"[{WorkbookIndex}]{FirstSheet}:{LastSheet}!{Area.GetDisplayString()}";
    }
}