using System.Globalization;

namespace ClosedXML.Parser.Tests;

public readonly record struct ScalarValue(string Type, object Value)
{
    public ScalarValue(double value) : this("Number", value) { }

    public ScalarValue(bool value) : this("Logical", value) { }

    public ScalarValue(string value) : this("Text", value) { }
};

public record AstNode
{
    internal AstNode[] Children { get; init; } = Array.Empty<AstNode>();

    public virtual bool Equals(AstNode? other) => other is not null && Children.SequenceEqual(other.Children);

    public override int GetHashCode() => Children.Sum(child => child.GetHashCode());
};

internal record ValueNode(string Type, object Value) : AstNode;

internal record ArrayNode(int Rows, int Columns, IReadOnlyList<ScalarValue> Elements) : AstNode
{
    public virtual bool Equals(ArrayNode? other)
    {
        if (ReferenceEquals(null, other))
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return base.Equals(other) &&
               Rows == other.Rows &&
               Columns == other.Columns &&
               Elements.SequenceEqual(other.Elements);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            int hashCode = base.GetHashCode();
            hashCode = (hashCode * 397) ^ Rows;
            hashCode = (hashCode * 397) ^ Columns;
            return hashCode;
        }
    }
}

internal record LocalReferenceNode(CellArea Reference) : AstNode;

internal record NameNode(string Name) : AstNode;

internal record ExternalReferenceNode(int WorkbookIndex, CellArea Reference) : AstNode;

internal record FunctionNode(string? Sheet, string Name) : AstNode
{
    public FunctionNode(string name) : this(null, name)
    {
    }
};

internal record ExternalFunctionNode(int WorkbookIndex, string? Sheet, string Name) : AstNode;

internal record StructureReferenceNode(string? Table, StructuredReferenceArea Area, string? FirstColumn, string? LastColumn) : AstNode;

internal record ExternalStructureReferenceNode(int WorkbookIndex, string Table, StructuredReferenceArea Area, string? FirstColumn, string? LastColumn) : AstNode;

internal record UnaryNode(UnaryOperation Operation) : AstNode;

internal record BinaryNode(BinaryOperation Operation) : AstNode
{
    public BinaryNode(BinaryOperation operation, AstNode left, AstNode right)
        : this(operation)
    {
        Children = new AstNode[] { left, right };
    }

    public AstNode Left => Children[0];

    public AstNode Right => Children[1];
};

#nullable disable

internal class F : IAstFactory<ScalarValue, AstNode>
{
    public ScalarValue LogicalValue(bool value)
    {
        return new ScalarValue("Logical", value);
    }

    public ScalarValue NumberValue(double value)
    {
        return new ScalarValue("Number", value);
    }

    public ScalarValue TextValue(ReadOnlySpan<char> input)
    {
        return new ScalarValue("Text", input.ToString());
    }

    public ScalarValue ErrorValue(ReadOnlySpan<char> error)
    {
        return new ScalarValue("Error", error.ToString());
    }

    public AstNode BlankNode()
    {
        return new ValueNode("Blank", string.Empty);
    }

    public AstNode LogicalNode(bool value)
    {
        return new ValueNode("Logical", value);
    }

    public AstNode ErrorNode(ReadOnlySpan<char> error)
    {
        return new ValueNode("Error", error.ToString());
    }

    public AstNode NumberNode(double value)
    {
        return new ValueNode("Number", value);
    }

    public AstNode TextNode(ReadOnlySpan<char> text)
    {
        return new ValueNode("Text", text.ToString());
    }

    public AstNode ArrayNode(int rows, int columns, IReadOnlyList<ScalarValue> array)
    {
        return new ArrayNode(rows, columns, array);
    }

    public AstNode LocalCellReference(ReadOnlySpan<char> input, CellArea area)
    {
        return new LocalReferenceNode(area);
    }

    public AstNode ExternalCellReference(ReadOnlySpan<char> input, int workbookIndex, CellArea area)
    {
        return new ExternalReferenceNode(workbookIndex, area);
    }

    public AstNode Function(ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(null, name.ToString()) { Children = args.ToArray() };
    }

    public AstNode Function(ReadOnlySpan<char> sheetName, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(sheetName.ToString(), name.ToString()) { Children = args.ToArray() };
    }

    public AstNode ExternalFunction(int workbookIndex, ReadOnlySpan<char> sheetName, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new ExternalFunctionNode(workbookIndex, sheetName.ToString(), name.ToString())
        {
            Children = args.ToArray()
        };
    }

    public AstNode ExternalFunction(int workbookIndex, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new ExternalFunctionNode(workbookIndex, null, name.ToString())
        {
            Children = args.ToArray()
        };
    }

    public AstNode StructureReference(ReadOnlySpan<char> text, StructuredReferenceArea area, string firstColumn, string lastColumn)
    {
        return new StructureReferenceNode(null, area, firstColumn, lastColumn);
    }

    public AstNode StructureReference(ReadOnlySpan<char> text, string table, StructuredReferenceArea area, string firstColumn, string lastColumn)
    {
        return new StructureReferenceNode(table, area, firstColumn, lastColumn);
    }

    public AstNode ExternalStructureReference(ReadOnlySpan<char> text, int workbookIndex, string table, StructuredReferenceArea area, string firstColumn, string lastColumn)
    {
        return new ExternalStructureReferenceNode(workbookIndex, table, area, firstColumn, lastColumn);
    }

    public AstNode LocalNameReference(ReadOnlySpan<char> name)
    {
        return new NameNode(name.ToString());
    }

    public AstNode LocalNameReference(ReadOnlySpan<char> sheet, ReadOnlySpan<char> name)
    {
        return default;
    }

    public AstNode ExternalNameReference(int workbookIndex, ReadOnlySpan<char> name)
    {
        return default;
    }

    public AstNode BinaryNode(BinaryOperation operation, AstNode leftNode, AstNode rightNode)
    {
        return new BinaryNode(operation, leftNode, rightNode);
    }

    public AstNode Unary(UnaryOperation operation, AstNode node)
    {
        return new UnaryNode(operation) { Children = new[] { node } };
    }
}