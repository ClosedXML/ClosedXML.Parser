namespace ClosedXML.Parser;

public record Ctx;

public class F : IAstFactory<ScalarValue, AstNode, Ctx>
{
    public ScalarValue LogicalValue(Ctx _, bool value)
    {
        return new ScalarValue("Logical", value);
    }

    public ScalarValue NumberValue(Ctx _, double value)
    {
        return new ScalarValue("Number", value);
    }

    public ScalarValue TextValue(Ctx _, ReadOnlySpan<char> input)
    {
        return new ScalarValue("Text", input.ToString());
    }

    public ScalarValue ErrorValue(Ctx _, ReadOnlySpan<char> error)
    {
        return new ScalarValue("Error", error.ToString());
    }

    public AstNode BlankNode(Ctx _)
    {
        return new ValueNode("Blank", string.Empty);
    }

    public AstNode LogicalNode(Ctx _, bool value)
    {
        return new ValueNode("Logical", value);
    }

    public AstNode ErrorNode(Ctx _, ReadOnlySpan<char> error)
    {
        return new ValueNode("Error", error.ToString());
    }

    public AstNode NumberNode(Ctx _, double value)
    {
        return new ValueNode("Number", value);
    }

    public AstNode TextNode(Ctx _, ReadOnlySpan<char> text)
    {
        return new ValueNode("Text", text.ToString());
    }

    public AstNode ArrayNode(Ctx _, int rows, int columns, IReadOnlyList<ScalarValue> array)
    {
        return new ArrayNode(rows, columns, array);
    }

    public AstNode Reference(Ctx _, ReferenceArea area)
    {
        return new ReferenceNode(area);
    }

    public AstNode SheetReference(Ctx _, string sheet, ReferenceArea area)
    {
        return new SheetReferenceNode(sheet, area);
    }

    public AstNode Reference3D(Ctx _, string firstSheet, string lastSheet, ReferenceArea area)
    {
        return new Reference3DNode(firstSheet, lastSheet, area);
    }

    public AstNode ExternalSheetReference(Ctx _, int workbookIndex, string sheet, ReferenceArea area)
    {
        return new ExternalSheetReferenceNode(workbookIndex, sheet, area);
    }

    public AstNode ExternalReference3D(Ctx _, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea area)
    {
        return new ExternalReference3DNode(workbookIndex, firstSheet, lastSheet, area);
    }

    public AstNode Function(Ctx _, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(null, name.ToString()) { Children = args.ToArray() };
    }

    public AstNode Function(Ctx _, string sheet, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(sheet, name.ToString()) { Children = args.ToArray() };
    }

    public AstNode ExternalFunction(Ctx _, int workbookIndex, string sheet, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new ExternalFunctionNode(workbookIndex, sheet, name.ToString())
        {
            Children = args.ToArray()
        };
    }

    public AstNode ExternalFunction(Ctx _, int workbookIndex, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new ExternalFunctionNode(workbookIndex, null, name.ToString())
        {
            Children = args.ToArray()
        };
    }

    public AstNode StructureReference(Ctx _, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return new StructureReferenceNode(null, area, firstColumn, lastColumn);
    }

    public AstNode StructureReference(Ctx _, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return new StructureReferenceNode(table, area, firstColumn, lastColumn);
    }

    public AstNode ExternalStructureReference(Ctx _, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return new ExternalStructureReferenceNode(workbookIndex, table, area, firstColumn, lastColumn);
    }

    public AstNode CellFunction(Ctx _, Reference cell, IReadOnlyList<AstNode> args)
    {
        return new CellFunctionNode(cell)
        {
            Children = args.ToArray()
        };
    }

    public AstNode Name(Ctx _, string name)
    {
        return new NameNode(name);
    }

    public AstNode SheetName(Ctx _, string sheet, string name)
    {
        return new SheetNameNode(sheet, name);
    }

    public AstNode ExternalName(Ctx _, int workbookIndex, string name)
    {
        return new ExternalNameNode(workbookIndex, name);
    }

    public AstNode ExternalSheetName(Ctx _, int workbookIndex, string sheet, string name)
    {
        return new ExternalSheetNameNode(workbookIndex, sheet, name);
    }

    public AstNode BinaryNode(Ctx _, BinaryOperation operation, AstNode leftNode, AstNode rightNode)
    {
        return new BinaryNode(operation, leftNode, rightNode);
    }

    public AstNode Unary(Ctx _, UnaryOperation operation, AstNode node)
    {
        return new UnaryNode(operation) { Children = new[] { node } };
    }

    public AstNode Nested(Ctx _, AstNode node)
    {
        // AST is mostly there for testing and visualization, so omit the unnecessary nodes in a tree.
        return node;
    }
}