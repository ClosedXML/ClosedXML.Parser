namespace ClosedXML.Parser;

#nullable disable

public class F : IAstFactory<ScalarValue, AstNode>
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

    public AstNode ArrayNode(NodeRange range, int rows, int columns, IReadOnlyList<ScalarValue> array)
    {
        return new ArrayNode(rows, columns, array);
    }

    public AstNode Reference(ReadOnlySpan<char> input, CellArea area)
    {
        return new ReferenceNode(area);
    }

    public AstNode ExternalReference(ReadOnlySpan<char> input, int workbookIndex, CellArea area)
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

    public AstNode CellFunction(Reference cell, IReadOnlyList<AstNode> args)
    {
        return new CellFunctionNode(cell)
        {
            Children = args.ToArray()
        };
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