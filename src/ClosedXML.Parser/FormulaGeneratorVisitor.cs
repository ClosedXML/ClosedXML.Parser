using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// A visitor that generates the identical formula for the parsed formula. It's
/// designed to allow modifications of references, e.g. renaming, moving references
/// and so on. Just inherit it and override one of <c>virtual Modify*</c> methods.
/// </summary>
internal class FormulaGeneratorVisitor : IAstFactory<TransformedSymbol, TransformedSymbol, TransformContext>
{
    // 1 quote on left, 1 quote on right size and at most 4 quotes inside.
    private const int QUOTE_RESERVE = 6;
    private const int SHEET_SEPARATOR_LEN = 1;
    private const int BOOK_PREFIX_LEN = 3;
    private const int MAX_R1_C1_LEN = 20;

    public virtual TransformedSymbol LogicalValue(TransformContext ctx, SymbolRange range, bool value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol NumberValue(TransformContext ctx, SymbolRange range, double value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol TextValue(TransformContext ctx, SymbolRange range, string text)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol ErrorValue(TransformContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol ArrayNode(TransformContext ctx, SymbolRange range, int rows, int columns, IReadOnlyList<TransformedSymbol> elements)
    {
        var sb = new StringBuilder(2 + elements.Sum(x => x.Length) + elements.Count);
        sb.Append('{');
        var i = 0;
        sb.Append(elements[i++].AsSpan());
        for (var col = 1; col < columns; ++col)
            sb.Append(',').Append(elements[i++].AsSpan());

        for (var row = 1; row < rows; ++row)
        {
            sb.Append(';');
            sb.Append(elements[i++].AsSpan());
            for (var col = 1; col < columns; ++col)
                sb.Append(',').Append(elements[i++].AsSpan());
        }

        sb.Append('}');
        var nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol BlankNode(TransformContext ctx, SymbolRange range)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol LogicalNode(TransformContext ctx, SymbolRange range, bool value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol ErrorNode(TransformContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol NumberNode(TransformContext ctx, SymbolRange range, double value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol TextNode(TransformContext ctx, SymbolRange range, string text)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol Reference(TransformContext ctx, SymbolRange range, ReferenceArea reference)
    {
        var sb = new StringBuilder(MAX_R1_C1_LEN);
        var transformedReference = ModifyRef(ctx, reference);
        var nodeText = sb.AppendRef(transformedReference).ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol SheetReference(TransformContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        var nodeText = sb
            .AppendSheetName(ModifySheet(ctx, sheet))
            .AppendReferenceSeparator()
            .AppendRef(ModifyRef(ctx, reference))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol Reference3D(TransformContext ctx, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        firstSheet = ModifySheet(ctx, firstSheet);
        lastSheet = ModifySheet(ctx, lastSheet);
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
            .AppendRef(ModifyRef(ctx, reference))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol ExternalSheetReference(TransformContext ctx, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        var nodeText = sb
            .AppendExternalSheetName(workbookIndex, ModifySheet(ctx, sheet))
            .AppendReferenceSeparator()
            .AppendRef(ModifyRef(ctx, reference))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol ExternalReference3D(TransformContext ctx, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        firstSheet = ModifySheet(ctx, firstSheet);
        lastSheet = ModifySheet(ctx, lastSheet);
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
            .AppendRef(ModifyRef(ctx, reference))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol Function(TransformContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var transformedFunction = ModifyFunction(ctx, functionName);
        var nodeText = sb.AppendFunction(transformedFunction, arguments).ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol ExternalFunction(TransformContext ctx, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var nodeText = sb
            .AppendBookIndex(workbookIndex)
            .AppendReferenceSeparator()
            .AppendFunction(ModifyFunction(ctx, functionName), arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol Function(TransformContext ctx, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var nodeText = sb
            .AppendSheetName(ModifySheet(ctx, sheetName))
            .AppendReferenceSeparator()
            .AppendFunction(ModifyFunction(ctx, functionName), arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol ExternalFunction(TransformContext ctx, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        var nodeText = sb
            .AppendExternalSheetName(workbookIndex, ModifySheet(ctx, sheetName))
            .AppendReferenceSeparator()
            .AppendFunction(ModifyFunction(ctx, functionName), arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol CellFunction(TransformContext ctx, SymbolRange range, RowCol cell, IReadOnlyList<TransformedSymbol> arguments)
    {
        var sb = new StringBuilder(MAX_R1_C1_LEN + SHEET_SEPARATOR_LEN + arguments.Sum(static x => x.Length));
        var nodeText = sb
            .AppendRef(ModifyCellFunction(ctx, cell))
            .AppendArguments(arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol StructureReference(TransformContext ctx, SymbolRange range,
        StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        var nodeText = GetIntraTableReference(area, firstColumn, lastColumn);
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol StructureReference(TransformContext ctx, SymbolRange range,
        string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        var nodeText = table + GetIntraTableReference(area, firstColumn, lastColumn);
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol ExternalStructureReference(TransformContext ctx,
        SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn,
        string? lastColumn)
    {
        var nodeText = new StringBuilder()
            .AppendBookIndex(workbookIndex).Append(table)
            .Append(GetIntraTableReference(area, firstColumn, lastColumn))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol Name(TransformContext ctx, SymbolRange range, string name)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    public virtual TransformedSymbol SheetName(TransformContext ctx, SymbolRange range, string sheet, string name)
    {
        var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        var nodeText = sb
            .AppendSheetName(ModifySheet(ctx, sheet))
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol ExternalName(TransformContext ctx, SymbolRange range, int workbookIndex, string name)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + SHEET_SEPARATOR_LEN + name.Length);
        var nodeText = sb
            .AppendBookIndex(workbookIndex)
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol ExternalSheetName(TransformContext ctx, SymbolRange range, int workbookIndex, string sheet, string name)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        var nodeText = sb
            .AppendExternalSheetName(workbookIndex, ModifySheet(ctx, sheet))
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol BinaryNode(TransformContext ctx, SymbolRange range, BinaryOperation operation, TransformedSymbol leftNode, TransformedSymbol rightNode)
    {
        var operand = operation switch
        {
            BinaryOperation.Concat => "&",
            BinaryOperation.Addition => "+",
            BinaryOperation.Subtraction => "-",
            BinaryOperation.Multiplication => "*",
            BinaryOperation.Division => "/",
            BinaryOperation.Power => "^",
            BinaryOperation.GreaterOrEqualThan => ">=",
            BinaryOperation.LessOrEqualThan => "<=",
            BinaryOperation.LessThan => "<",
            BinaryOperation.GreaterThan => ">",
            BinaryOperation.NotEqual => "<>",
            BinaryOperation.Equal => "=",
            BinaryOperation.Union => ",",
            BinaryOperation.Intersection => " ",
            BinaryOperation.Range => ":",
            _ => throw new NotSupportedException()
        };
        var nodeText = new StringBuilder(leftNode.Length + operand.Length + rightNode.Length)
            .Append(leftNode.AsSpan())
            .Append(operand)
            .Append(rightNode.AsSpan())
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol Unary(TransformContext ctx, SymbolRange range, UnaryOperation operation, TransformedSymbol node)
    {
        var (isPrefix, opChar) = operation switch
        {
            UnaryOperation.Plus => (true, '+'),
            UnaryOperation.Minus => (true, '-'),
            UnaryOperation.Percent => (false, '%'),
            UnaryOperation.ImplicitIntersection => (true, '@'),
            UnaryOperation.SpillRange => (false, '#'),
            _ => throw new NotSupportedException()
        };
        var sb = new StringBuilder(node.Length + 1);
        if (isPrefix)
            sb.Append(opChar).Append(node.AsSpan());
        else
            sb.Append(node.AsSpan()).Append(opChar);

        var nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    public virtual TransformedSymbol Nested(TransformContext ctx, SymbolRange range, TransformedSymbol node)
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

    /// <summary>
    /// Modify reference to a cell.
    /// </summary>
    /// <param name="ctx">The origin of formula.</param>
    /// <param name="reference">Area reference.</param>
    /// <returns>Modified reference.</returns>
    protected virtual ReferenceArea ModifyRef(TransformContext ctx, ReferenceArea reference)
    {
        return reference;
    }

    /// <summary>
    /// Modify reference to a cell function.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="cell">Original cell containing function.</param>
    /// <returns>Modified reference.</returns>
    protected virtual RowCol ModifyCellFunction(TransformContext ctx, RowCol cell)
    {
        return cell;
    }

    /// <summary>
    /// An extension to modify sheet name, e.g. rename.
    /// </summary>
    /// <param name="ctx">The transformation context.</param>
    /// <param name="sheetName">Original sheet name.</param>
    /// <returns>New sheet name.</returns>
    protected virtual string ModifySheet(TransformContext ctx, string sheetName)
    {
        return sheetName;
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
}