namespace ClosedXML.Parser.Ast;

public record FunctionNode(string? Sheet, string Name) : AstNode
{
    public FunctionNode(string name) : this(null, name)
    {
    }
};