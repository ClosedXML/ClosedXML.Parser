using System;
using System.Collections.Generic;

namespace ClosedXML.Parser;

/// <summary>
/// A factory used to create an AST through <see cref="FormulaParser{TScalarValue,TNode}"/>.
/// </summary>
/// <typeparam name="TScalarValue">Type of a scalar value used across expressions.</typeparam>
/// <typeparam name="TNode">Type of a node used in the AST.</typeparam>
public interface IAstFactory<TScalarValue, TNode>
    where TNode : class
{
    /// <summary>
    /// Create a logical value for an array item.
    /// </summary>
    /// <param name="value">The logical value of an array.</param>
    TScalarValue LogicalValue(bool value);

    /// <summary>
    /// Create a numerical value for an array item.
    /// </summary>
    /// <param name="value">The numeric value of an array. Never <c>NaN</c> or <c>Infinity</c>.</param>
    TScalarValue NumberValue(double value);

    /// <summary>
    /// Create a text value for an array item.
    /// </summary>
    /// <param name="text">The text. The characters of text are already unescaped.</param>
    TScalarValue TextValue(ReadOnlySpan<char> text);

    TScalarValue ErrorValue(ReadOnlySpan<char> error);

    TNode ArrayNode(int rows, int columns, IList<TScalarValue> elements);

    TNode BlankNode();

    TNode LogicalNode(bool value);

    TNode ErrorNode(ReadOnlySpan<char> error);

    TNode NumberNode(double value);

    TNode TextNode(ReadOnlySpan<char> text);

    TNode LocalCellReference(ReadOnlySpan<char> input, CellArea area);

    TNode ExternalCellReference(ReadOnlySpan<char> input, int workbookIndex, CellArea area);

    TNode Function(ReadOnlySpan<char> name, IList<TNode> args);

    TNode StructureReference(ReadOnlySpan<char> text, StructuredReferenceSpecific specific, string firstColumn, string lastColumn);

    TNode StructureReference(ReadOnlySpan<char> text, string table, StructuredReferenceSpecific specific, string firstColumn, string lastColumn);

    TNode ExternalStructureReference(ReadOnlySpan<char> text, int workbookIndex, string table, StructuredReferenceSpecific specific, string firstColumn, string lastColumn);

    TNode LocalNameReference(ReadOnlySpan<char> name);

    TNode LocalNameReference(ReadOnlySpan<char> sheet, ReadOnlySpan<char> name);

    TNode ExternalNameReference(int workbookIndex, ReadOnlySpan<char> name);

    TNode BinaryNode(BinaryOperation operation, TNode leftNode, TNode rightNode);

    TNode Unary(UnaryOperation operation, TNode node);
}
