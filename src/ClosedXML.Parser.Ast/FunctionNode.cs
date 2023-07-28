namespace ClosedXML.Parser;

public record FunctionNode(string? Sheet, string Name) : AstNode
{
    public FunctionNode(string name) : this(null, name)
    {
    }

    public override string GetDisplayString(ReferenceStyle style)
    {
        return Sheet is not null ? $"{Sheet}!{Name}" : Name;
    }
}