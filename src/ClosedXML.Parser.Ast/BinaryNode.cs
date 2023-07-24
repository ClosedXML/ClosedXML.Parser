namespace ClosedXML.Parser.Ast;

public record BinaryNode(BinaryOperation Operation) : AstNode
{
    public BinaryNode(BinaryOperation operation, AstNode left, AstNode right)
        : this(operation)
    {
        Children = new AstNode[] { left, right };
    }

    public AstNode Left => Children[0];

    public AstNode Right => Children[1];
};