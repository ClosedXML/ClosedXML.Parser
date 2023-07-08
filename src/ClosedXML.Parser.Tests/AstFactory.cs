#nullable disable

using System.Globalization;

namespace ClosedXML.Parser.Tests;

public readonly record struct ScalarValue(string Type, object Value);

internal record AstNode
{
    internal AstNode[] Children { get; init; } = Array.Empty<AstNode>();
};

internal record ValueNode(string Type, object Value) : AstNode;

internal record LocalReferenceNode(CellArea Reference) : AstNode;

internal record ExternalReferenceNode(int WorkbookIndex, CellArea Reference) : AstNode;

internal class F : IAstFactory<ScalarValue, AstNode>
{
    public ScalarValue LogicalValue(bool value)
    {
        return new ScalarValue("Logical", value);
    }

    public ScalarValue NumberValue(double value)
    {
        return new ScalarValue("Number", value.ToString(CultureInfo.InvariantCulture));
    }

    public ScalarValue TextValue(ReadOnlySpan<char> input)
    {
        return new ScalarValue("Text", input.ToString());
    }

    public ScalarValue ErrorValue(string input, int firstIndex, int length)
    {
        return new ScalarValue("Error", input.Substring(firstIndex, length));
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

    public AstNode ArrayNode(int rows, int columns, IList<ScalarValue> array)
    {
        return new ValueNode("Array", $"{{{rows}x{columns}}}");
    }

    public AstNode LocalCellReference(ReadOnlySpan<char> input, CellArea area)
    {
        return new LocalReferenceNode(area);
    }

    public AstNode ExternalCellReference(ReadOnlySpan<char> input, int workbookIndex, CellArea area)
    {
        return new ExternalReferenceNode(workbookIndex, area);
    }

    public AstNode Function(ReadOnlySpan<char> name, IList<AstNode> args)
    {
        return default;
    }

    public AstNode Function(ReadOnlySpan<char> name, AstNode[] args)
    {
        return default;
    }

    public AstNode StructureReference(ReadOnlySpan<char> intraTableReference)
    {
        return default;
    }

    public AstNode StructureReference(ReadOnlySpan<char> tableName, ReadOnlySpan<char> intraTableReference)
    {
        return default;
    }

    public AstNode StructureReference(ReadOnlySpan<char> bookPrefix, ReadOnlySpan<char> tableName, ReadOnlySpan<char> intraTableReference)
    {
        return default;
    }

    public AstNode LocalNameReference(ReadOnlySpan<char> name)
    {
        return default;
    }

    public AstNode LocalNameReference(ReadOnlySpan<char> sheet, ReadOnlySpan<char> name)
    {
        return default;
    }

    public AstNode ExternalNameReference(ReadOnlySpan<char> bookPrefix, ReadOnlySpan<char> name)
    {
        return default;
    }

    public AstNode BinaryNode(BinaryOperation operation, AstNode leftNode, AstNode rightNode)
    {
        return new ValueNode("Binary", operation) { Children = new[] { leftNode, rightNode } };
    }

    public AstNode Unary(UnaryOperation operation, AstNode node)
    {
        return default;
    }
}