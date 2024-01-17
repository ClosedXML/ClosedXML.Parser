using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// It's designed to allow modifications of references, e.g. renaming, moving references
/// and so on. Just inherit it and override one of <c>virtual Modify*</c> methods.
/// </summary>
public class RefModVisitor : IAstFactory<TransformedSymbol, TransformedSymbol, ModContext>
{
    private const string REF_ERROR = "#REF!";
    private static readonly CopyVisitor s_copyVisitor = new();

    /// <inheritdoc />
    public TransformedSymbol LogicalValue(ModContext ctx, SymbolRange range, bool value)
    {
        return s_copyVisitor.LogicalValue(ctx, range, value);
    }

    /// <inheritdoc />
    public TransformedSymbol NumberValue(ModContext ctx, SymbolRange range, double value)
    {
        return s_copyVisitor.NumberValue(ctx, range, value);
    }

    /// <inheritdoc />
    public TransformedSymbol TextValue(ModContext ctx, SymbolRange range, string text)
    {
        return s_copyVisitor.TextValue(ctx, range, text);
    }

    /// <inheritdoc />
    public TransformedSymbol ErrorValue(ModContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return s_copyVisitor.ErrorValue(ctx, range, error);
    }

    /// <inheritdoc />
    public TransformedSymbol ArrayNode(ModContext ctx, SymbolRange range, int rows, int columns, IReadOnlyList<TransformedSymbol> elements)
    {
        return s_copyVisitor.ArrayNode(ctx, range, rows, columns, elements);
    }

    /// <inheritdoc />
    public TransformedSymbol BlankNode(ModContext ctx, SymbolRange range)
    {
        return s_copyVisitor.BlankNode(ctx, range);
    }

    /// <inheritdoc />
    public TransformedSymbol LogicalNode(ModContext ctx, SymbolRange range, bool value)
    {
        return s_copyVisitor.LogicalNode(ctx, range, value);
    }

    /// <inheritdoc />
    public TransformedSymbol ErrorNode(ModContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        if (range.Length == error.Length)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Deal with `Sheet!REF!`, `#REF!A1` and `#REF!#REF!`
        var symbol = ctx.Formula.AsSpan().Slice(range.Start, range.Length);
        var sheetIsRefError = symbol.StartsWith(REF_ERROR.AsSpan(), StringComparison.OrdinalIgnoreCase);

        if (sheetIsRefError)
        {
            // #REF!A1 is invalid formula that can't be parsed by Excel. It is displayed, but
            // likely only because it is a serialization of internal structures. When sheet is
            // deleted, the result is #REF!, which is how it is actually saved in the file.
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);
        }

        // Sheet!#REF! is a valid formula per grammar and Excel, though it displeases me.
        var sheet = symbol.Slice(0, symbol.Length - REF_ERROR.Length - 1);
        var modifiedSheet = ModifySheet(ctx, sheet.ToString());
        var nodeText = new StringBuilder()
            .AppendSheetReference(modifiedSheet)
            .Append(symbol.Slice(symbol.Length - REF_ERROR.Length))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public TransformedSymbol NumberNode(ModContext ctx, SymbolRange range, double value)
    {
        return s_copyVisitor.NumberNode(ctx, range, value);
    }

    /// <inheritdoc />
    public TransformedSymbol TextNode(ModContext ctx, SymbolRange range, string text)
    {
        return s_copyVisitor.TextNode(ctx, range, text);
    }

    /// <inheritdoc />
    public TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.Reference(ctx, range, modifiedReference.Value);
    }

    /// <inheritdoc />
    public TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        var modifiedSheet = ModifySheet(ctx, sheet);
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedSheet is null || modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.SheetReference(ctx, range, modifiedSheet, modifiedReference.Value);
    }

    /// <inheritdoc />
    public TransformedSymbol Reference3D(ModContext ctx, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var modifiedFirstSheet = ModifySheet(ctx, firstSheet);
        var modifiedLastSheet = ModifySheet(ctx, lastSheet);
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedFirstSheet is null || modifiedLastSheet is null || modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.Reference3D(ctx, range, modifiedFirstSheet, modifiedLastSheet, modifiedReference.Value);
    }

    /// <inheritdoc />
    public TransformedSymbol ExternalSheetReference(ModContext ctx, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.ExternalSheetReference(ctx, range, workbookIndex, sheet, modifiedReference.Value);
    }

    /// <inheritdoc />
    public TransformedSymbol ExternalReference3D(ModContext ctx, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var modifiedFirstSheet = ModifySheet(ctx, firstSheet);
        var modifiedLastSheet = ModifySheet(ctx, lastSheet);
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedFirstSheet is null || modifiedLastSheet is null || modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.ExternalReference3D(ctx, range, workbookIndex, modifiedFirstSheet, modifiedLastSheet, modifiedReference.Value);
    }

    /// <inheritdoc />
    public TransformedSymbol Function(ModContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        return s_copyVisitor.Function(ctx, range, functionName, arguments);
    }

    /// <inheritdoc />
    public TransformedSymbol Function(ModContext ctx, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var modifiedFunction = ModifyFunction(ctx, functionName);
        var modifiedSheet = ModifySheet(ctx, sheetName);
        if (modifiedSheet is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.Function(ctx, range, modifiedSheet, modifiedFunction, arguments);
    }

    /// <inheritdoc />
    public TransformedSymbol ExternalFunction(ModContext ctx, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        return s_copyVisitor.ExternalFunction(ctx, range, workbookIndex, sheetName, functionName, arguments);
    }

    /// <inheritdoc />
    public TransformedSymbol ExternalFunction(ModContext ctx, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        return s_copyVisitor.ExternalFunction(ctx, range, workbookIndex, functionName, arguments);
    }

    /// <inheritdoc />
    public TransformedSymbol CellFunction(ModContext ctx, SymbolRange range, RowCol cell, IReadOnlyList<TransformedSymbol> arguments)
    {
        var modifiedCell = ModifyCellFunction(ctx, cell);
        if (modifiedCell is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.CellFunction(ctx, range, modifiedCell.Value, arguments);
    }

    /// <inheritdoc />
    public TransformedSymbol StructureReference(ModContext ctx, SymbolRange range, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return s_copyVisitor.StructureReference(ctx, range, area, firstColumn, lastColumn);
    }

    /// <inheritdoc />
    public TransformedSymbol StructureReference(ModContext ctx, SymbolRange range, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        var modifiedTableName = ModifyTable(ctx, table);
        if (modifiedTableName is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.StructureReference(ctx, range, modifiedTableName, area, firstColumn, lastColumn);
    }

    /// <inheritdoc />
    public TransformedSymbol ExternalStructureReference(ModContext ctx, SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return s_copyVisitor.ExternalStructureReference(ctx, range, workbookIndex, table, area, firstColumn, lastColumn);
    }

    /// <inheritdoc />
    public TransformedSymbol Name(ModContext ctx, SymbolRange range, string name)
    {
        return s_copyVisitor.Name(ctx, range, name);
    }

    /// <inheritdoc />
    public TransformedSymbol SheetName(ModContext ctx, SymbolRange range, string sheet, string name)
    {
        var modifiedSheet = ModifySheet(ctx, sheet);
        if (modifiedSheet is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.SheetName(ctx, range, modifiedSheet, name);
    }

    /// <inheritdoc />
    public TransformedSymbol ExternalName(ModContext ctx, SymbolRange range, int workbookIndex, string name)
    {
        return s_copyVisitor.ExternalName(ctx, range, workbookIndex, name);
    }

    /// <inheritdoc />
    public TransformedSymbol ExternalSheetName(ModContext ctx, SymbolRange range, int workbookIndex, string sheet, string name)
    {
        return s_copyVisitor.ExternalSheetName(ctx, range, workbookIndex, sheet, name);
    }

    /// <inheritdoc />
    public TransformedSymbol BinaryNode(ModContext ctx, SymbolRange range, BinaryOperation operation, TransformedSymbol leftNode, TransformedSymbol rightNode)
    {
        return s_copyVisitor.BinaryNode(ctx, range, operation, leftNode, rightNode);
    }

    /// <inheritdoc />
    public TransformedSymbol Unary(ModContext ctx, SymbolRange range, UnaryOperation operation, TransformedSymbol node)
    {
        return s_copyVisitor.Unary(ctx, range, operation, node);
    }

    /// <inheritdoc />
    public TransformedSymbol Nested(ModContext ctx, SymbolRange range, TransformedSymbol node)
    {
        return s_copyVisitor.Nested(ctx, range, node);
    }

    /// <summary>
    /// An extension to modify sheet name, e.g. rename.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="sheetName">Original sheet name.</param>
    /// <returns>New sheet name. If null, it indicates sheet has been deleted and should be replaced with <c>#REF!</c></returns>
    protected virtual string? ModifySheet(ModContext ctx, string sheetName)
    {
        return sheetName;
    }

    /// <summary>
    /// Modify reference to a cell.
    /// </summary>
    /// <param name="ctx">The origin of formula.</param>
    /// <param name="table">Original name of a table.</param>
    /// <returns>Modified name of a table or null if <c>#REF!</c>.</returns>
    protected virtual string? ModifyTable(ModContext ctx, string table)
    {
        return table;
    }

    /// <summary>
    /// An extension to modify name of a function.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="functionName">Original name of function.</param>
    /// <returns>New name of a function.</returns>
    protected virtual ReadOnlySpan<char> ModifyFunction(ModContext ctx, ReadOnlySpan<char> functionName)
    {
        return functionName;
    }

    /// <summary>
    /// Modify reference to a cell. This method is called for every place where is a ref and is
    /// mostly intended to change reference style.
    /// </summary>
    /// <param name="ctx">The origin of formula.</param>
    /// <param name="reference">Area reference.</param>
    /// <returns>Modified reference or null if <c>#REF!</c>.</returns>
    internal virtual ReferenceArea? ModifyRef(ModContext ctx, ReferenceArea reference)
    {
        return reference;
    }

    /// <summary>
    /// Modify reference to a cell function.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="cell">Original cell containing function.</param>
    /// <returns>Modified reference or null if <c>#REF!</c>.</returns>
    internal virtual RowCol? ModifyCellFunction(ModContext ctx, RowCol cell)
    {
        return cell;
    }
}
