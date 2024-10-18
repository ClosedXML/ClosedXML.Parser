using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// A visitor that generates the identical formula for the parsed formula based on passed arguments.
/// </summary>
public class CopyVisitor : IAstFactory<TransformedSymbol, TransformedSymbol, ModContext>
{
    // 1 quote on left, 1 quote on right size and at most 4 quotes inside.
    private const int QUOTE_RESERVE = 6;
    private const int SHEET_SEPARATOR_LEN = 1;
    private const int BOOK_PREFIX_LEN = 3;
    private const int MAX_R1_C1_LEN = 20;

    /// <inheritdoc />
    public virtual TransformedSymbol LogicalValue(ModContext ctx, SymbolRange range, bool value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol NumberValue(ModContext ctx, SymbolRange range, double value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol TextValue(ModContext ctx, SymbolRange range, string text)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ErrorValue(ModContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ArrayNode(ModContext ctx, SymbolRange range, int rows, int columns, IReadOnlyList<TransformedSymbol> elements)
    {
        var sb = new StringBuilder(2 + elements.Sum(x => x.Length) + elements.Count);
        sb.AppendStartFragment(ctx, range, elements[0]);
        var i = 0;
        sb.Append(elements[i++].AsSpan());
        for (var col = 1; col < columns; ++col)
        {
            sb.AppendMiddleFragment(ctx, elements[i - 1], elements[i]);
            sb.Append(elements[i++].AsSpan());
        }

        for (var row = 1; row < rows; ++row)
        {
            sb.AppendMiddleFragment(ctx, elements[i - 1], elements[i]);
            sb.Append(elements[i++].AsSpan());
            for (var col = 1; col < columns; ++col)
            {
                sb.AppendMiddleFragment(ctx, elements[i - 1], elements[i]);
                sb.Append(elements[i++].AsSpan());
            }
        }

        sb.AppendEndFragment(ctx, range, elements[elements.Count - 1]);
        var nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol BlankNode(ModContext ctx, SymbolRange range)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol LogicalNode(ModContext ctx, SymbolRange range, bool value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ErrorNode(ModContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol NumberNode(ModContext ctx, SymbolRange range, double value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol TextNode(ModContext ctx, SymbolRange range, string text)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        var sb = new StringBuilder(MAX_R1_C1_LEN);
        var nodeText = sb.AppendRef(reference).ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        var nodeText = sb
            .AppendSheetReference(sheet)
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol BangReference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        var sb = new StringBuilder(SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        var nodeText = sb
            .AppendReferenceSeparator()
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol Reference3D(ModContext ctx, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        if (NameUtils.ShouldQuote(firstSheet.AsSpan()) || NameUtils.ShouldQuote(lastSheet.AsSpan()))
        {
            sb
                .Append('\'')
                .AppendEscapedSheetName(firstSheet)
                .Append(':')
                .AppendEscapedSheetName(lastSheet)
                .Append('\'');
        }
        else
        {
            sb.Append(firstSheet)
                .Append(':')
                .Append(lastSheet);
        }

        var nodeText = sb
            .AppendReferenceSeparator()
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalSheetReference(ModContext ctx, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        var nodeText = sb
            .AppendExternalSheetReference(workbookIndex, sheet)
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalReference3D(ModContext ctx, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        if (NameUtils.ShouldQuote(firstSheet.AsSpan()) || NameUtils.ShouldQuote(lastSheet.AsSpan()))
        {
            sb
                .Append('\'')
                .AppendBookIndex(workbookIndex)
                .AppendEscapedSheetName(firstSheet)
                .Append(':')
                .AppendEscapedSheetName(lastSheet)
                .Append('\'');
        }
        else
        {
            sb
                .AppendBookIndex(workbookIndex)
                .Append(firstSheet)
                .Append(':')
                .Append(lastSheet);
        }

        var nodeText = sb
            .AppendReferenceSeparator()
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Function(ModContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var nodeText = sb.Append(functionName).AppendArguments(ctx, range, arguments).ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalFunction(ModContext ctx, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var nodeText = sb
            .AppendBookIndex(workbookIndex)
            .AppendReferenceSeparator()
            .AppendFunction(ctx, range, functionName, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Function(ModContext ctx, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var nodeText = sb
            .AppendSheetReference(sheetName)
            .AppendFunction(ctx, range, functionName, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalFunction(ModContext ctx, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var nodeText = sb
            .AppendExternalSheetReference(workbookIndex, sheetName)
            .AppendFunction(ctx, range, functionName, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol CellFunction(ModContext ctx, SymbolRange range, RowCol cell, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(MAX_R1_C1_LEN + SHEET_SEPARATOR_LEN + arguments.Sum(static x => x.Length));
        var nodeText = sb
            .AppendRef(cell)
            .AppendArguments(ctx, range, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol StructureReference(ModContext ctx, SymbolRange range,
        StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        var nodeText = GetIntraTableReference(area, firstColumn, lastColumn);
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol StructureReference(ModContext ctx, SymbolRange range,
        string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        var nodeText = table + GetIntraTableReference(area, firstColumn, lastColumn);
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalStructureReference(ModContext ctx,
        SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn,
        string? lastColumn)
    {
        var nodeText = new StringBuilder()
            .AppendBookIndex(workbookIndex).Append(table)
            .Append(GetIntraTableReference(area, firstColumn, lastColumn))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Name(ModContext ctx, SymbolRange range, string name)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol SheetName(ModContext ctx, SymbolRange range, string sheet, string name)
    {
        var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        var nodeText = sb
            .AppendSheetReference(sheet)
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalName(ModContext ctx, SymbolRange range, int workbookIndex, string name)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + SHEET_SEPARATOR_LEN + name.Length);
        var nodeText = sb
            .AppendBookIndex(workbookIndex)
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalSheetName(ModContext ctx, SymbolRange range, int workbookIndex, string sheet, string name)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        var nodeText = sb
            .AppendExternalSheetReference(workbookIndex, sheet)
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol BinaryNode(ModContext ctx, SymbolRange range, BinaryOperation operation, TransformedSymbol leftNode, TransformedSymbol rightNode)
    {
        var sb = new StringBuilder(leftNode.Length + rightNode.OriginalRange.Start - leftNode.OriginalRange.End + rightNode.Length)
            .AppendStartFragment(ctx, range, leftNode)
            .Append(leftNode.AsSpan())
            .AppendMiddleFragment(ctx, leftNode, rightNode)
            .Append(rightNode.AsSpan())
            .AppendEndFragment(ctx, range, rightNode);

        var nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Unary(ModContext ctx, SymbolRange range, UnaryOperation operation, TransformedSymbol node)
    {
        var sb = new StringBuilder(node.Length + 1)
            .AppendStartFragment(ctx, range, node)
            .Append(node.AsSpan())
            .AppendEndFragment(ctx, range, node);

        var nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Nested(ModContext ctx, SymbolRange range, TransformedSymbol node)
    {
        var nodeText = new StringBuilder(node.Length + 2)
            .AppendStartFragment(ctx, range, node)
            .Append(node.AsSpan())
            .AppendEndFragment(ctx, range, node)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    private static string GetIntraTableReference(StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        if (firstColumn is null || lastColumn is null)
        {
            // No column

            // Shorthand for full table inside the table.
            if (area == StructuredReferenceArea.None)
                return "[]";

            if (area == (StructuredReferenceArea.Headers | StructuredReferenceArea.Data))
                return "[[#Headers],[#Data]]";

            if (area == (StructuredReferenceArea.Data | StructuredReferenceArea.Totals))
                return "[[#Data],[#Totals]]";

            return Keyword(area);
        }

        if (firstColumn == lastColumn)
        {
            // One column
            if (area == StructuredReferenceArea.None)
            {
                // One column, no keyword
                return new StringBuilder(firstColumn.Length + 2)
                    .Append('[').Append(firstColumn).Append(']')
                    .ToString();
            }

            // One column, keyword
            var keywordList = KeywordList(area);
            return new StringBuilder(keywordList.Length + firstColumn.Length + 5)
                .Append('[')
                .Append(keywordList).Append(',')
                .Append('[').Append(firstColumn).Append(']')
                .Append(']')
                .ToString();
        }
        else
        {
            // Two columns
            var keywordList = KeywordList(area);
            var sb = new StringBuilder(firstColumn.Length + lastColumn.Length + keywordList.Length + 8);
            sb.Append('[');
            if (keywordList.Length > 0)
                sb.Append(keywordList).Append(',');

            return sb
                .Append('[').Append(firstColumn).Append(']')
                .Append(':')
                .Append('[').Append(lastColumn).Append(']')
                .Append(']')
                .ToString();
        }

        static string KeywordList(StructuredReferenceArea area)
        {
            return area switch
            {
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data => "[#Headers],[#Data]",
                StructuredReferenceArea.Data | StructuredReferenceArea.Totals => "[#Data],[#Totals]",
                _ => Keyword(area),
            };
        }

        static string Keyword(StructuredReferenceArea area)
        {
            return area switch
            {
                StructuredReferenceArea.None => string.Empty,
                StructuredReferenceArea.Headers => "[#Headers]",
                StructuredReferenceArea.Data => "[#Data]",
                StructuredReferenceArea.Totals => "[#Totals]",
                StructuredReferenceArea.All => "[#All]",
                StructuredReferenceArea.ThisRow => "[#This Row]",
                _ => throw new NotSupportedException(),
            };
        }
    }
}