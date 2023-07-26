namespace ClosedXML.Parser;

public record NameNode(string Name) : AstNode
{
    public override string GetDisplayString()
    {
        return Name;
    }
}