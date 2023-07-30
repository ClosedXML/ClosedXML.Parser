using System;
using System.Collections.Generic;

namespace ClosedXML.Parser;

/// <summary>
/// A factory used to create an AST through <see cref="FormulaParser{TScalarValue,TNode}"/>.
/// </summary>
/// <remarks>
/// Sheet names are in most cases strings, while most other texts are <c>ReadOnlySpan&lt;char&gt;</c>.
/// The reason is that sheet name is always used as-is and in some methods is null, while
/// other parameters might be processed. E.g. errors are likely to be transformed to enum, function name
/// might need conversion from text <c>IFS</c> to <c>_xlfn.IFS</c> and so on.
/// </remarks>
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

    /// <summary>
    /// Create an error for an array item.
    /// </summary>
    /// <param name="error">The error text, string with <c>#</c> until the end of an error. No whitespace, converted to upper case, no matter the input..</param>
    TScalarValue ErrorValue(ReadOnlySpan<char> error);

    /// <summary>
    /// Create an array for scalar values.
    /// </summary>
    /// <param name="rows">Number of rows of an array. At least 1.</param>
    /// <param name="columns">Number of column of an array. At least 1.</param>
    /// <param name="elements">Elements of an array, row by row. The number of elements is <paramref name="rows"/>*<paramref name="columns"/>.</param>
    TNode ArrayNode(int rows, int columns, IReadOnlyList<TScalarValue> elements);

    /// <summary>
    /// Create a blank node. In most cases, a blank argument of a function, e.g. <c>IF(TRUE,,)</c>.
    /// </summary>
    TNode BlankNode();

    /// <summary>
    /// Create a node with a logical value.
    /// </summary>
    TNode LogicalNode(bool value);

    /// <summary>
    /// Create a node with an error value.
    /// </summary>
    /// <param name="error">The error text, string with <c>#</c> until the end of an error. No whitespace. In upper case format.</param>
    TNode ErrorNode(ReadOnlySpan<char> error);

    /// <summary>
    /// Create a node with an error value.
    /// </summary>
    /// <param name="value">The numeric value of an array. Never <c>NaN</c> or <c>Infinity</c>.</param>
    TNode NumberNode(double value);

    /// <summary>
    /// Create a node with a text value.
    /// </summary>
    /// <param name="text">The text. The characters of text are already unescaped.</param>
    TNode TextNode(ReadOnlySpan<char> text);

    /// <summary>
    /// Create a node for a reference to cells without a worksheet.
    /// </summary>
    /// <param name="area">The referenced area.</param>
    TNode Reference(ReferenceArea area);

    /// <summary>
    /// Create a node for a reference in a specific sheet.
    /// </summary>
    /// <param name="sheet">Name of a sheet (unescaped) of the <paramref name="area"/>.</param>
    /// <param name="area">Area in the sheet.</param>
    TNode SheetReference(string sheet, ReferenceArea area);

    /// <summary>
    /// Create a node for a 3D reference.
    /// </summary>
    /// <param name="firstSheet">First sheet of 3D reference.</param>
    /// <param name="lastSheet">Last sheet of 3D reference.</param>
    /// <param name="area">Area in all sheets of 3D reference.</param>
    TNode Reference3D(string firstSheet, string lastSheet, ReferenceArea area);

    /// <summary>
    /// Create a node for a reference to cells of a specific sheet in a different worksheet.
    /// </summary>
    /// <param name="workbookIndex">Id of an external workbook. The actual path to the file is in workbook part, <c>externalReferences</c> tag.</param>
    /// <param name="sheet">Name of a sheet (unescaped) of the <paramref name="area"/>.</param>
    /// <param name="area">Area the external sheet.</param>
    TNode ExternalSheetReference(int workbookIndex, string sheet, ReferenceArea area);

    /// <summary>
    /// Create a node for a 3D reference in a different workbook.
    /// </summary>
    /// <param name="workbookIndex">Id of an external workbook. The actual path to the file is in workbook part, <c>externalReferences</c> tag.</param>
    /// <param name="firstSheet">First sheet of 3D reference.</param>
    /// <param name="lastSheet">Last sheet of 3D reference.</param>
    /// <param name="area">Area in all sheets of 3D reference.</param>
    TNode ExternalReference3D(int workbookIndex, string firstSheet, string lastSheet, ReferenceArea area);

    /// <summary>
    /// Create a node for a function.
    /// </summary>
    /// <param name="functionName">Name of a function.</param>
    /// <param name="arguments">Nodes of argument values.</param>
    TNode Function(ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments);

    /// <summary>
    /// Create a node for a function on a sheet. Might happen for VBA.
    /// </summary>
    /// <param name="sheetName">Name of a sheet.</param>
    /// <param name="functionName">Name of a function.</param>
    /// <param name="args">Nodes of argument values.</param>
    TNode Function(string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> args);

    TNode ExternalFunction(int workbookIndex, string sheet, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments);

    TNode ExternalFunction(int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments);

    /// <summary>
    /// Create a cell function. It references another function that should likely contain a LAMBDA value.
    /// </summary>
    /// <remarks>Cell functions are not yet supported by Excel, but are part of a grammar.</remarks>
    /// <param name="cell">A reference to a cell with a LAMBDA.</param>
    /// <param name="arguments">Arguments to pass to a LAMBDA.</param>
    TNode CellFunction(Reference cell, IReadOnlyList<TNode> arguments);

    /// <summary>
    /// Create a node to represent a structure reference without a table to a range of columns.
    /// Such reference is only allowed in the table (e.g. total formulas).
    /// </summary>
    /// <param name="area">A portion of a table that should be considered.</param>
    /// <param name="firstColumn">The first column of a range. Null, if whole table. If only one column, same as <paramref name="lastColumn"/>.</param>
    /// <param name="lastColumn">The last column of a range. Null, if whole table.If only one column, same as <paramref name="firstColumn"/>.</param>
    TNode StructureReference(StructuredReferenceArea area, string? firstColumn, string? lastColumn);

    /// <summary>
    /// Create a node to represent a structure reference to a table.
    /// </summary>
    /// <param name="table">A name of a table.</param>
    /// <param name="area">A portion of a table that should be considered.</param>
    /// <param name="firstColumn">The first column of a range. Null, if whole table. If only one column, same as <paramref name="lastColumn"/>.</param>
    /// <param name="lastColumn">The last column of a range. Null, if whole table.If only one column, same as <paramref name="firstColumn"/>.</param>
    TNode StructureReference(string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn);

    /// <summary>
    /// Create a node to represent a structure reference to a table in some other workbook.
    /// </summary>
    /// <param name="workbookIndex">Id of external workbook.</param>
    /// <param name="table">A name of a table.</param>
    /// <param name="area">A portion of a table that should be considered.</param>
    /// <param name="firstColumn">The first column of a range. Null, if whole table. If only one column, same as <paramref name="lastColumn"/>.</param>
    /// <param name="lastColumn">The last column of a range. Null, if whole table.If only one column, same as <paramref name="firstColumn"/>.</param>
    TNode ExternalStructureReference(int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn);

    /// <summary>
    /// Create a node that should evaluate to a value of a name defined in a workbook.
    /// </summary>
    /// <remarks>
    /// Name can be any formula, though in most cases, it is a cell reference. Also note that
    /// names can be global (usable in a whole workbook) or local (only for one worksheet).
    /// </remarks>
    /// <param name="name">The defined name.</param>
    TNode Name(string name);

    /// <summary>
    /// Create a node that should evaluate to a value of a name defined in a worksheet.
    /// </summary>
    /// <remarks>
    /// Name can be any formula, though in most cases, it is a cell reference. Also note that
    /// names can be global (usable in a whole workbook) or local (only for one worksheet).
    /// </remarks>
    /// <param name="sheet">Name of a sheet, unescaped.</param>
    /// <param name="name">The defined name.</param>
    TNode SheetName(string sheet, string name);

    /// <summary>
    /// Create a node that should evaluate to a value of a defined name in a different workbook.
    /// </summary>
    /// <param name="workbookIndex">Id of an external workbook. The actual path to the file is in workbook part, <c>externalReferences</c> tag.</param>
    /// <param name="name">Name from a workbook. It can be defined name or a name of a table.</param>
    TNode ExternalName(int workbookIndex, string name);

    /// <summary>
    /// Create a node that should evaluate to a value of a defined name in a worksheet of a different workbook.
    /// </summary>
    /// <remarks>
    /// Name can be any formula, though in most cases, it is a cell reference. Also note that
    /// names can be global (usable in a whole workbook) or local (only for one worksheet).
    /// </remarks>
    /// <param name="workbookIndex">Id of an external workbook. The actual path to the file is in workbook part, <c>externalReferences</c> tag.</param>
    /// <param name="sheet">Name of a sheet in the external workbook, unescaped.</param>
    /// <param name="name">The defined name.</param>
    TNode ExternalSheetName(int workbookIndex, string sheet, string name);

    /// <summary>
    /// Create a node that performs a binary operation on values from another nodes.
    /// </summary>
    /// <param name="operation">Binary operation.</param>
    /// <param name="leftNode">Node that should be evaluated for left argument of a binary operation.</param>
    /// <param name="rightNode">Node that should be evaluated for right argument of a binary operation.</param>
    TNode BinaryNode(BinaryOperation operation, TNode leftNode, TNode rightNode);

    /// <summary>
    /// Create a node that performs an unary operation on a value from another node.
    /// </summary>
    /// <param name="operation">Unary operation.</param>
    /// <param name="node">Node that should be evaluated for a value.</param>
    TNode Unary(UnaryOperation operation, TNode node);

    /// <summary>
    /// This factory method is called for nested expression in braces (<c>(1+2)/4</c>). The problem isn't that
    /// it would be evaluated incorrectly, but it is to preserve braces during A1 to R1C1 transformation.
    /// </summary>
    /// <remarks>Simplest implementation returns the same node and avoids extra nodes.</remarks>
    /// <param name="node">The node representing expression in braces.</param>
    TNode Nested(TNode node);
}
