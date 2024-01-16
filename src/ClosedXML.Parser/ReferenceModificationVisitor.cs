using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// It's designed to allow modifications of references, e.g. renaming, moving references
/// and so on. Just inherit it and override one of <c>virtual Modify*</c> methods.
/// </summary>
internal class ReferenceModificationVisitor : IAstFactory<TransformedSymbol, TransformedSymbol, TransformContext>
{
    private const string REF_ERROR = "#REF!";
    private static readonly CopyVisitor s_copyVisitor = new();

    public TransformedSymbol LogicalValue(TransformContext ctx, SymbolRange range, bool value)
    {
        return s_copyVisitor.LogicalValue(ctx, range, value);
    }

    public TransformedSymbol NumberValue(TransformContext ctx, SymbolRange range, double value)
    {
        return s_copyVisitor.NumberValue(ctx, range, value);
    }

    public TransformedSymbol TextValue(TransformContext ctx, SymbolRange range, string text)
    {
        return s_copyVisitor.TextValue(ctx, range, text);
    }

    public TransformedSymbol ErrorValue(TransformContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return s_copyVisitor.ErrorValue(ctx, range, error);
    }

    public TransformedSymbol ArrayNode(TransformContext ctx, SymbolRange range, int rows, int columns, IReadOnlyList<TransformedSymbol> elements)
    {
        return s_copyVisitor.ArrayNode(ctx, range, rows, columns, elements);
    }

    public TransformedSymbol BlankNode(TransformContext ctx, SymbolRange range)
    {
        return s_copyVisitor.BlankNode(ctx, range);
    }

    public TransformedSymbol LogicalNode(TransformContext ctx, SymbolRange range, bool value)
    {
        return s_copyVisitor.LogicalNode(ctx, range, value);
    }

    public TransformedSymbol ErrorNode(TransformContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        if (range.Length == error.Length)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Deal with `Sheet!REF!`, `#REF!A1` and `#REF!#REF!`
        var symbol = ctx.Formula.AsSpan().Slice(range.Start, range.Length);
        var sheetIsRefError = symbol.StartsWith(REF_ERROR.AsSpan(), StringComparison.OrdinalIgnoreCase);
        var referenceIsRefError = symbol.EndsWith(REF_ERROR.AsSpan(), StringComparison.OrdinalIgnoreCase);

        if (sheetIsRefError && referenceIsRefError)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        if (sheetIsRefError)
        {
            var referenceSymbol = symbol.Slice(REF_ERROR.Length);
            var reference = TokenParser.ParseReference(referenceSymbol, ctx.IsA1);
            var modifiedReference = ModifyRef(ctx, reference);
            var nodeText = new StringBuilder()
                .Append(symbol.Slice(0, REF_ERROR.Length))
                .AppendRef(modifiedReference)
                .ToString();
            return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
        }
        else
        {
            var sheet = symbol.Slice(0, symbol.Length - REF_ERROR.Length - 1);
            var modifiedSheet = ModifySheet(ctx, sheet.ToString());
            var nodeText = new StringBuilder()
                .AppendSheetReference(modifiedSheet)
                .Append(symbol.Slice(symbol.Length - REF_ERROR.Length))
                .ToString();
            return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
        }
    }

    public TransformedSymbol NumberNode(TransformContext ctx, SymbolRange range, double value)
    {
        return s_copyVisitor.NumberNode(ctx, range, value);
    }

    public TransformedSymbol TextNode(TransformContext ctx, SymbolRange range, string text)
    {
        return s_copyVisitor.TextNode(ctx, range, text);
    }

    public TransformedSymbol Reference(TransformContext ctx, SymbolRange range, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.Reference(ctx, range, modifiedReference.Value);
    }

    public TransformedSymbol SheetReference(TransformContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        var modifiedSheet = ModifySheet(ctx, sheet);
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedSheet is null || modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.SheetReference(ctx, range, modifiedSheet, modifiedReference.Value);
    }

    public TransformedSymbol Reference3D(TransformContext ctx, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var modifiedFirstSheet = ModifySheet(ctx, firstSheet);
        var modifiedLastSheet = ModifySheet(ctx, lastSheet);
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedFirstSheet is null || modifiedLastSheet is null || modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.Reference3D(ctx, range, modifiedFirstSheet, modifiedLastSheet, modifiedReference.Value);
    }

    public TransformedSymbol ExternalSheetReference(TransformContext ctx, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
    {
        var modifiedSheet = ModifyExternalSheet(ctx, workbookIndex, sheet);
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedSheet is null || modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.ExternalSheetReference(ctx, range, workbookIndex, modifiedSheet, modifiedReference.Value);
    }

    public TransformedSymbol ExternalReference3D(TransformContext ctx, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var modifiedFirstSheet = ModifySheet(ctx, firstSheet);
        var modifiedLastSheet = ModifySheet(ctx, lastSheet);
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedFirstSheet is null || modifiedLastSheet is null || modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.ExternalReference3D(ctx, range, workbookIndex, modifiedFirstSheet, modifiedLastSheet, modifiedReference.Value);
    }

    public TransformedSymbol Function(TransformContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        return s_copyVisitor.Function(ctx, range, functionName, arguments);
    }

    public TransformedSymbol Function(TransformContext ctx, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var modifiedFunction = ModifyFunction(ctx, functionName);
        var modifiedSheet = ModifySheet(ctx, sheetName);
        if (modifiedSheet is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.Function(ctx, range, modifiedSheet, modifiedFunction, arguments);
    }

    public TransformedSymbol ExternalFunction(TransformContext ctx, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var modifiedFunction = ModifyFunction(ctx, functionName);
        var modifiedSheetName = ModifyExternalSheet(ctx, workbookIndex, sheetName);
        if (modifiedSheetName is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.ExternalFunction(ctx, range, workbookIndex, modifiedSheetName, modifiedFunction, arguments);
    }

    public TransformedSymbol ExternalFunction(TransformContext ctx, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        return s_copyVisitor.ExternalFunction(ctx, range, workbookIndex, functionName, arguments);
    }

    public TransformedSymbol CellFunction(TransformContext ctx, SymbolRange range, RowCol cell, IReadOnlyList<TransformedSymbol> arguments)
    {
        var modifiedCell = ModifyCellFunction(ctx, cell);
        if (modifiedCell is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.CellFunction(ctx, range, modifiedCell.Value, arguments);
    }

    public TransformedSymbol StructureReference(TransformContext ctx, SymbolRange range, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return s_copyVisitor.StructureReference(ctx, range, area, firstColumn, lastColumn);
    }

    public TransformedSymbol StructureReference(TransformContext ctx, SymbolRange range, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return s_copyVisitor.StructureReference(ctx, range, table, area, firstColumn, lastColumn);
    }

    public TransformedSymbol ExternalStructureReference(TransformContext ctx, SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return s_copyVisitor.ExternalStructureReference(ctx, range, workbookIndex, table, area, firstColumn, lastColumn);
    }

    public TransformedSymbol Name(TransformContext ctx, SymbolRange range, string name)
    {
        return s_copyVisitor.Name(ctx, range, name);
    }

    public TransformedSymbol SheetName(TransformContext ctx, SymbolRange range, string sheet, string name)
    {
        var modifiedSheet = ModifySheet(ctx, sheet);
        if (modifiedSheet is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.SheetName(ctx, range, modifiedSheet, name);
    }

    public TransformedSymbol ExternalName(TransformContext ctx, SymbolRange range, int workbookIndex, string name)
    {
        return s_copyVisitor.ExternalName(ctx, range, workbookIndex, name);
    }

    public TransformedSymbol ExternalSheetName(TransformContext ctx, SymbolRange range, int workbookIndex, string sheet, string name)
    {
        var modifiedSheet = ModifyExternalSheet(ctx, workbookIndex, sheet);
        if (modifiedSheet is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return s_copyVisitor.ExternalSheetName(ctx, range, workbookIndex, modifiedSheet, name);
    }

    public TransformedSymbol BinaryNode(TransformContext ctx, SymbolRange range, BinaryOperation operation, TransformedSymbol leftNode, TransformedSymbol rightNode)
    {
        return s_copyVisitor.BinaryNode(ctx, range, operation, leftNode, rightNode);
    }

    public TransformedSymbol Unary(TransformContext ctx, SymbolRange range, UnaryOperation operation, TransformedSymbol node)
    {
        return s_copyVisitor.Unary(ctx, range, operation, node);
    }

    public TransformedSymbol Nested(TransformContext ctx, SymbolRange range, TransformedSymbol node)
    {
        return s_copyVisitor.Nested(ctx, range, node);
    }

    /// <summary>
    /// An extension to modify sheet name, e.g. rename.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="sheetName">Original sheet name.</param>
    /// <returns>New sheet name. If null, it indicates sheet has been deleted and should be replaced with <c>#REF!</c></returns>
    protected virtual string? ModifySheet(TransformContext ctx, string sheetName)
    {
        return sheetName;
    }

    protected virtual string? ModifyExternalSheet(TransformContext ctx, int bookIndex, string sheetName)
    {
        return sheetName;
    }

    /// <summary>
    /// Modify reference to a cell.
    /// </summary>
    /// <param name="ctx">The origin of formula.</param>
    /// <param name="reference">Area reference.</param>
    /// <returns>Modified reference or null if <c>#REF!</c>.</returns>
    protected virtual ReferenceArea? ModifyRef(TransformContext ctx, ReferenceArea reference)
    {
        return reference;
    }

    /// <summary>
    /// An extension to modify name of a function.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="functionName">Original name of function.</param>
    /// <returns>New name of a function.</returns>
    protected virtual ReadOnlySpan<char> ModifyFunction(TransformContext ctx, ReadOnlySpan<char> functionName)
    {
        return functionName;
    }

    /// <summary>
    /// Modify reference to a cell function.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="cell">Original cell containing function.</param>
    /// <returns>Modified reference or null if <c>#REF!</c>.</returns>
    protected virtual RowCol? ModifyCellFunction(TransformContext ctx, RowCol cell)
    {
        return cell;
    }
}
