using System;
using System.Collections.Generic;

namespace ClosedXML.Lexer;

/// <summary>
/// A factory used to create an AST through <see cref="FormulaParser{TScalarValue,TNode}"/>.
/// </summary>
/// <typeparam name="TScalarValue">Type of a scalar value used across expressions.</typeparam>
/// <typeparam name="TNode">Type of a node used in the AST.</typeparam>
public interface IAstFactory<TScalarValue, TNode>
    where TNode : class
{
    TScalarValue LogicalValue(bool value);

    TScalarValue NumberValue(double value);

    TScalarValue TextValue(ReadOnlySpan<char> text);

    TScalarValue ErrorValue(string input, int firstIndex, int length);

    TNode ArrayNode(int rows, int column, IList<TScalarValue> elements);

    TNode BlankNode();

    TNode LogicalNode(bool value);

    TNode ErrorNode(string input, int firstIndex, int length);

    TNode NumberNode(double value);

    TNode TextNode(ReadOnlySpan<char> text);

    TNode LocalCellReference(string input, int firstIndex, int length);

    TNode ExternalCellReference(string input, int firstIndex, int length);

    TNode Function(ReadOnlySpan<char> name, IList<TNode> args);

    TNode StructureReference(ReadOnlySpan<char> intraTableReference);

    TNode StructureReference(ReadOnlySpan<char> tableName, ReadOnlySpan<char> intraTableReference);

    TNode StructureReference(ReadOnlySpan<char> bookPrefix, ReadOnlySpan<char> tableName, ReadOnlySpan<char> intraTableReference);

    TNode LocalNameReference(ReadOnlySpan<char> name);

    TNode LocalNameReference(ReadOnlySpan<char> sheet, ReadOnlySpan<char> name);

    TNode ExternalNameReference(ReadOnlySpan<char> bookPrefix, ReadOnlySpan<char> name);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="operation">One of <c>'≥'</c>, <c>'≤'</c>, <c>'&lt;'</c>, <c>'&gt;'</c>, <c>'≠'</c>, <c>'='</c>, <c>'&amp;'</c>, <c>'+'</c>, <c>'-'</c>, <c>'^'</c>, <c>','</c>, <c>' '</c>, <c>':'</c></param>
    /// <param name="leftNode"></param>
    /// <param name="rightNode"></param>
    /// <returns></returns>
    TNode BinaryNode(BinaryOperation operation, TNode leftNode, TNode rightNode);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="operation"><c>+</c>, <c>-</c> or <c>%</c>.</param>
    /// <param name="node"></param>
    /// <returns></returns>
    TNode Unary(char operation, TNode node);
}
