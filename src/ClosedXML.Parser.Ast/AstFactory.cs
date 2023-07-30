﻿namespace ClosedXML.Parser;

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

    public AstNode ArrayNode(int rows, int columns, IReadOnlyList<ScalarValue> array)
    {
        return new ArrayNode(rows, columns, array);
    }

    public AstNode Reference(ReferenceArea area)
    {
        return new ReferenceNode(area);
    }

    public AstNode SheetReference(string sheet, ReferenceArea area)
    {
        return new SheetReferenceNode(sheet, area);
    }

    public AstNode Reference3D(string firstSheet, string lastSheet, ReferenceArea area)
    {
        return new Reference3DNode(firstSheet, lastSheet, area);
    }

    public AstNode ExternalSheetReference(int workbookIndex, string sheet, ReferenceArea area)
    {
        return new ExternalSheetReferenceNode(workbookIndex, sheet, area);
    }

    public AstNode ExternalReference3D(int workbookIndex, string firstSheet, string lastSheet, ReferenceArea area)
    {
        return new ExternalReference3DNode(workbookIndex, firstSheet, lastSheet, area);
    }

    public AstNode Function(ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(null, name.ToString()) { Children = args.ToArray() };
    }

    public AstNode Function(string sheet, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new FunctionNode(sheet, name.ToString()) { Children = args.ToArray() };
    }

    public AstNode ExternalFunction(int workbookIndex, string sheet, ReadOnlySpan<char> name, IReadOnlyList<AstNode> args)
    {
        return new ExternalFunctionNode(workbookIndex, sheet, name.ToString())
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

    public AstNode StructureReference(StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return new StructureReferenceNode(null, area, firstColumn, lastColumn);
    }

    public AstNode StructureReference(string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return new StructureReferenceNode(table, area, firstColumn, lastColumn);
    }

    public AstNode ExternalStructureReference(int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
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

    public AstNode Name(string name)
    {
        return new NameNode(name);
    }

    public AstNode SheetName(string sheet, string name)
    {
        return new SheetNameNode(sheet, name);
    }

    public AstNode ExternalName(int workbookIndex, string name)
    {
        return new ExternalNameNode(workbookIndex, name);
    }

    public AstNode ExternalSheetName(int workbookIndex, string sheet, string name)
    {
        return new ExternalSheetNameNode(workbookIndex, sheet, name);
    }

    public AstNode BinaryNode(BinaryOperation operation, AstNode leftNode, AstNode rightNode)
    {
        return new BinaryNode(operation, leftNode, rightNode);
    }

    public AstNode Unary(UnaryOperation operation, AstNode node)
    {
        return new UnaryNode(operation) { Children = new[] { node } };
    }

    public AstNode Nested(AstNode node)
    {
        // AST is mostly there for testing and visualization, so omit the unnecessary nodes in a tree.
        return node;
    }
}