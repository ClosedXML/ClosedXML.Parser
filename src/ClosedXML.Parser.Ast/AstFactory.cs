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

    public ScalarValue TextValue(Ctx _, string input)
    {
        return new ScalarValue("Text", input);
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

    public AstNode TextNode(Ctx _, string text)
    {
        return new ValueNode("Text", text);
    }

    public AstNode ArrayNode(Ctx _, int rows, int columns, IReadOnlyList<ScalarValue> array)
    {
        return new ArrayNode(rows, columns, array);
    }

    public AstNode Reference(Ctx _, SymbolRange range, ReferenceArea reference)
    {
        return new ReferenceNode(reference);
    }

    public AstNode SheetReference(Ctx _, SymbolRange range, string sheet, ReferenceArea reference)
    {
        return new SheetReferenceNode(sheet, reference);
    }

    public AstNode Reference3D(Ctx _, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        return new Reference3DNode(firstSheet, lastSheet, reference);
    }

    public AstNode ExternalSheetReference(Ctx _, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
    {
        return new ExternalSheetReferenceNode(workbookIndex, sheet, reference);
    }

    public AstNode ExternalReference3D(Ctx _, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        return new ExternalReference3DNode(workbookIndex, firstSheet, lastSheet, reference);
    }

    public AstNode Function(Ctx _, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(null, name.ToString()) { Children = args.ToArray() };
    }

    public AstNode Function(Ctx _, string sheet, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(sheet, name.ToString()) { Children = args.ToArray() };
    }

    public AstNode ExternalFunction(Ctx _, int workbookIndex, string sheetName, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new ExternalFunctionNode(workbookIndex, sheetName, name.ToString())
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

    public AstNode CellFunction(Ctx _, RowCol cell, IReadOnlyList<AstNode> args)
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