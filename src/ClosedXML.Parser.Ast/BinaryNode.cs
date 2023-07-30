namespace ClosedXML.Parser;

public record BinaryNode(BinaryOperation Operation) : AstNode
{
    private static readonly Dictionary<BinaryOperation, string> OpNames = new()
    {
        { BinaryOperation.Concat, "&" },
        { BinaryOperation.GreaterOrEqualThan, ">=" },
        { BinaryOperation.LessOrEqualThan, "<=" },
        { BinaryOperation.LessThan, "<" },
        { BinaryOperation.GreaterThan, ">" },
        { BinaryOperation.NotEqual, "!=" },
        { BinaryOperation.Equal, "=" },
        { BinaryOperation.Addition, "+" },
        { BinaryOperation.Subtraction, "-" },
        { BinaryOperation.Multiplication, "*" },
        { BinaryOperation.Division, "/" },
        { BinaryOperation.Power, "^" },
        { BinaryOperation.Union, "union" },
        { BinaryOperation.Intersection, "intersection" },
        { BinaryOperation.Range, "range" },
    };

    public BinaryNode(BinaryOperation operation, AstNode left, AstNode right)
        : this(operation)
    {
        Children = new[] { left, right };
    }

    public override string GetDisplayString(ReferenceStyle style) => OpNames[Operation];
};