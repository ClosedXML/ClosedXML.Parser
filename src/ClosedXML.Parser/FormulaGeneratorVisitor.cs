using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// A visitor that generates the identical formula for the parsed formula. It's
/// designed to allow modifications of references, e.g. renaming, moving references
/// and so on. Just inherit it and override one of <c>virtual Modify*</c> methods.
/// </summary>
public class FormulaGeneratorVisitor : IAstFactory<string, string, (int Row, int Col)>
{
    // 1 quote on left, 1 quote on right size and at most 4 quotes inside.
    private const int QUOTE_RESERVE = 6;
    private const int SHEET_SEPARATOR_LEN = 1;
    private const int BOOK_PREFIX_LEN = 3;
    private const int MAX_R1_C1_LEN = 20;

    string IAstFactory<string, string, (int Row, int Col)>.LogicalValue((int, int) _, bool value)
    {
        return value ? "TRUE" : "FALSE";
    }

    string IAstFactory<string, string, (int Row, int Col)>.NumberValue((int, int) _, double value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }

    string IAstFactory<string, string, (int Row, int Col)>.TextValue((int, int) _, string text)
    {
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    string IAstFactory<string, string, (int Row, int Col)>.ErrorValue((int, int) _, ReadOnlySpan<char> error)
    {
        return error.ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.ArrayNode((int, int) _, int rows, int columns, IReadOnlyList<string> elements)
    {
        var sb = new StringBuilder(2 + elements.Sum(x => x.Length) + elements.Count);
        sb.Append('{');
        var i = 0;
        sb.Append(elements[i++]);
        for (var col = 1; col < columns; ++col)
            sb.Append(',').Append(elements[i++]);

        for (var row = 1; row < rows; ++row)
        {
            sb.Append(';');
            sb.Append(elements[i++]);
            for (var col = 1; col < columns; ++col)
                sb.Append(',').Append(elements[i++]);
        }

        sb.Append('}');
        return sb.ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.BlankNode((int, int) _)
    {
        return string.Empty;
    }

    string IAstFactory<string, string, (int Row, int Col)>.LogicalNode((int, int) _, bool value)
    {
        return value ? "TRUE" : "FALSE";
    }

    string IAstFactory<string, string, (int Row, int Col)>.ErrorNode((int, int) _, ReadOnlySpan<char> error)
    {
        return error.ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.NumberNode((int, int) _, double value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }

    string IAstFactory<string, string, (int Row, int Col)>.TextNode((int, int) _, string text)
    {
        return new StringBuilder(text.Length + QUOTE_RESERVE)
            .Append('"')
            .Append(text).Replace("\"", "\"\"", 1, text.Length)
            .Append('\"')
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.Reference((int Row, int Col) point, SymbolRange _, ReferenceArea reference)
    {
        var sb = new StringBuilder(MAX_R1_C1_LEN);
        return sb
            .AppendRef(ModifyRef(reference, point))
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.SheetReference((int Row, int Col) point, SymbolRange range, string sheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        return sb
            .AppendSheetName(ModifySheet(sheet))
            .AppendReferenceSeparator()
            .AppendRef(ModifyRef(reference, point))
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.Reference3D((int Row, int Col) point, SymbolRange _, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        firstSheet = ModifySheet(firstSheet);
        lastSheet = ModifySheet(lastSheet);
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

        return sb
            .AppendReferenceSeparator()
            .AppendRef(ModifyRef(reference, point))
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.ExternalSheetReference((int Row, int Col) point, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        return sb
            .AppendExternalSheetName(workbookIndex, ModifySheet(sheet))
            .AppendReferenceSeparator()
            .AppendRef(ModifyRef(reference, point))
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.ExternalReference3D((int Row, int Col) point, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        firstSheet = ModifySheet(firstSheet);
        lastSheet = ModifySheet(lastSheet);
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

        return sb
            .AppendReferenceSeparator()
            .AppendRef(ModifyRef(reference, point))
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.Function((int, int) _, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
    {
        var sb = new StringBuilder(functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        return sb.AppendFunction(ModifyFunction(functionName), arguments).ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.ExternalFunction((int, int) _, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        return sb
            .AppendBookIndex(workbookIndex)
            .AppendReferenceSeparator()
            .AppendFunction(ModifyFunction(functionName), arguments)
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.Function((int Row, int Col) point, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
    {
        var sb = new StringBuilder(sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        return sb
            .AppendSheetName(ModifySheet(sheetName))
            .AppendReferenceSeparator()
            .AppendFunction(ModifyFunction(functionName), arguments)
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.ExternalFunction((int Row, int Col) point, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        return sb
            .AppendExternalSheetName(workbookIndex, ModifySheet(sheetName))
            .AppendReferenceSeparator()
            .AppendFunction(ModifyFunction(functionName), arguments)
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.CellFunction((int Row, int Col) point, SymbolRange range, RowCol cell, IReadOnlyList<string> arguments)
    {
        var sb = new StringBuilder(MAX_R1_C1_LEN + SHEET_SEPARATOR_LEN + arguments.Sum(static x => x.Length));
        return sb
            .AppendRef(ModifyCellFunction(cell, point))
            .AppendArguments(arguments)
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.StructureReference((int, int) _, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return GetIntraTableReference(area, firstColumn, lastColumn);
    }

    string IAstFactory<string, string, (int Row, int Col)>.StructureReference((int, int) _, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return table + GetIntraTableReference(area, firstColumn, lastColumn);
    }

    string IAstFactory<string, string, (int Row, int Col)>.ExternalStructureReference((int, int) _, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        return new StringBuilder()
            .AppendBookIndex(workbookIndex).Append(table)
            .Append(GetIntraTableReference(area, firstColumn, lastColumn))
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.Name((int, int) _, string name)
    {
        return name;
    }

    string IAstFactory<string, string, (int Row, int Col)>.SheetName((int Row, int Col) point, string sheet, string name)
    {
        var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        return sb
            .AppendSheetName(ModifySheet(sheet))
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.ExternalName((int, int) _, int workbookIndex, string name)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + SHEET_SEPARATOR_LEN + name.Length);
        return sb
            .AppendBookIndex(workbookIndex)
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.ExternalSheetName((int Row, int Col) point, int workbookIndex, string sheet, string name)
    {
        var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        return sb
            .AppendExternalSheetName(workbookIndex, ModifySheet(sheet))
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
    }

    string IAstFactory<string, string, (int Row, int Col)>.BinaryNode((int, int) _, BinaryOperation operation, string leftNode, string rightNode)
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
        return leftNode + operand + rightNode;
    }

    string IAstFactory<string, string, (int Row, int Col)>.Unary((int, int) _, UnaryOperation operation, string node)
    {
        return operation switch
        {
            UnaryOperation.Plus => '+' + node,
            UnaryOperation.Minus => '-' + node,
            UnaryOperation.Percent => node + '%',
            UnaryOperation.ImplicitIntersection => '@' + node,
            UnaryOperation.SpillRange => node + '#',
            _ => throw new NotSupportedException()
        };
    }

    string IAstFactory<string, string, (int Row, int Col)>.Nested((int, int) _, string node)
    {
        return "(" + node + ")";
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
    /// <param name="reference">Area reference.</param>
    /// <param name="point">The origin of formula.</param>
    /// <returns>Modified reference.</returns>
    protected virtual ReferenceArea ModifyRef(ReferenceArea reference, (int Row, int Col) point)
    {
        return reference;
    }

    /// <summary>
    /// Modify reference to a cell function.
    /// </summary>
    /// <param name="cell">Original cell containing function.</param>
    /// <param name="point">The origin of formula.</param>
    /// <returns>Modified reference.</returns>
    protected virtual RowCol ModifyCellFunction(RowCol cell, (int Row, int Col) point)
    {
        return cell;
    }

    /// <summary>
    /// An extension to modify sheet name, e.g. rename.
    /// </summary>
    /// <param name="sheetName">Original sheet name.</param>
    /// <returns>New sheet name.</returns>
    protected virtual string ModifySheet(string sheetName)
    {
        return sheetName;
    }

    /// <summary>
    /// An extension to modify name of a function.
    /// </summary>
    /// <param name="functionName">Original name of function.</param>
    /// <returns>New name of a function.</returns>
    protected virtual ReadOnlySpan<char> ModifyFunction(ReadOnlySpan<char> functionName)
    {
        return functionName;
    }
}