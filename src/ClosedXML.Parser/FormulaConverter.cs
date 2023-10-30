using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// Convert between <em>A1</em> and <em>R1C1</em> style formulas.
/// </summary>
public static class FormulaConverter
{
    private static readonly TextVisitorR1C1 s_visitorR1C1 = new();

    /// <summary>
    /// Convert a formula in <em>A1</em> form to the <em>R1C1</em> form.
    /// </summary>
    /// <param name="formulaA1">Formula text.</param>
    /// <param name="row">The row origin of R1C1, from 1 to 1048576.</param>
    /// <param name="col">The column origin of R1C1, from 1 to 16384.</param>
    /// <returns>Formula converted to R1C1.</returns>
    /// <exception cref="ParsingException">The formula is not parseable.</exception>
    public static string ToR1C1(string formulaA1, int row, int col)
    {
        return FormulaParser<string, string, (int Row, int Col)>.CellFormulaA1(formulaA1, (row, col), s_visitorR1C1);
    }

    private class TextVisitorR1C1 : TextVisitor
    {
        protected override StringBuilder AppendRefAsText(ReferenceSymbol reference, (int Row, int Col) point, StringBuilder sb)
        {
            return reference.AppendR1C1(sb, point.Row, point.Col);
        }

        protected override StringBuilder AppendRefAsText(RowCol cell, (int Row, int Col) point, StringBuilder sb)
        {
            cell.AppendR1C1(sb, point.Row, point.Col);
            return sb;
        }
    }

    private abstract class TextVisitor : IAstFactory<string, string, (int Row, int Col)>
    {
        // 1 quote on left, 1 quote on right size and at most 4 quotes inside.
        private const int QUOTE_RESERVE = 6;
        private const int SHEET_SEPARATOR_LEN = 1;
        private const int BOOK_PREFIX_LEN = 3;
        private const int MAX_R1_C1_LEN = 20;

        public string LogicalValue((int, int) _, bool value)
        {
            return value ? "TRUE" : "FALSE";
        }

        public string NumberValue((int, int) _, double value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        public string TextValue((int, int) _, string text)
        {
            return "\"" + text.Replace("\"", "\"\"") + "\"";
        }

        public string ErrorValue((int, int) _, ReadOnlySpan<char> error)
        {
            return error.ToString();
        }

        public string ArrayNode((int, int) _, int rows, int columns, IReadOnlyList<string> elements)
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

        public string BlankNode((int, int) _)
        {
            return string.Empty;
        }

        public string LogicalNode((int, int) _, bool value)
        {
            return value ? "TRUE" : "FALSE";
        }

        public string ErrorNode((int, int) _, ReadOnlySpan<char> error)
        {
            return error.ToString();
        }

        public string NumberNode((int, int) _, double value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        public string TextNode((int, int) _, string text)
        {
            return new StringBuilder(text.Length + QUOTE_RESERVE)
                .Append('"')
                .Append(text).Replace("\"", "\"\"", 1, text.Length)
                .Append('\"')
                .ToString();
        }

        public string Reference((int Row, int Col) point, ReferenceSymbol reference)
        {
            var sb = new StringBuilder(MAX_R1_C1_LEN);
            return AppendRefAsText(reference, point, sb).ToString();
        }

        public string SheetReference((int Row, int Col) point, string sheet, ReferenceSymbol reference)
        {
            var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
            sb
                .AppendSheetName(sheet)
                .AppendReferenceSeparator();

            return AppendRefAsText(reference, point, sb).ToString();
        }

        public string Reference3D((int Row, int Col) point, string firstSheet, string lastSheet, ReferenceSymbol reference)
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

            sb.AppendReferenceSeparator();
            return AppendRefAsText(reference, point, sb).ToString();
        }

        public string ExternalSheetReference((int Row, int Col) point, int workbookIndex, string sheet, ReferenceSymbol reference)
        {
            var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
            sb
                .AppendExternalSheetName(workbookIndex, sheet)
                .AppendReferenceSeparator();
            return AppendRefAsText(reference, point, sb).ToString();
        }

        public string ExternalReference3D((int Row, int Col) point, int workbookIndex, string firstSheet, string lastSheet, ReferenceSymbol reference)
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

            sb.AppendReferenceSeparator();
            return AppendRefAsText(reference, point, sb).ToString();
        }

        public string Function((int, int) _, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            var sb = new StringBuilder(functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
            return sb.AppendFunction(functionName, arguments).ToString();
        }

        public string ExternalFunction((int, int) _, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            var sb = new StringBuilder(BOOK_PREFIX_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
            return sb
                .AppendBookIndex(workbookIndex)
                .AppendReferenceSeparator()
                .AppendFunction(functionName, arguments)
                .ToString();
        }

        public string Function((int, int) _, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            var sb = new StringBuilder(sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
            return sb
                .AppendSheetName(sheetName)
                .AppendReferenceSeparator()
                .AppendFunction(functionName, arguments)
                .ToString();
        }

        public string ExternalFunction((int, int) _, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            var sb = new StringBuilder(BOOK_PREFIX_LEN + sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
            return sb
                .AppendExternalSheetName(workbookIndex, sheetName)
                .AppendReferenceSeparator()
                .AppendFunction(functionName, arguments)
                .ToString();
        }

        public string CellFunction((int Row, int Col) point, RowCol cell, IReadOnlyList<string> arguments)
        {
            var sb = new StringBuilder(MAX_R1_C1_LEN + SHEET_SEPARATOR_LEN + arguments.Sum(static x => x.Length));
            return AppendRefAsText(cell, point, sb).AppendArguments(arguments).ToString();
        }

        public string StructureReference((int, int) _, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            return GetIntraTableReference(area, firstColumn, lastColumn);
        }

        public string StructureReference((int, int) _, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            return table + GetIntraTableReference(area, firstColumn, lastColumn);
        }

        public string ExternalStructureReference((int, int) _, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            return new StringBuilder()
                .AppendBookIndex(workbookIndex).Append(table)
                .Append(GetIntraTableReference(area, firstColumn, lastColumn))
                .ToString();
        }

        public string Name((int, int) _, string name)
        {
            return name;
        }

        public string SheetName((int, int) _, string sheet, string name)
        {
            var sb = new StringBuilder(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
            return sb
                .AppendSheetName(sheet)
                .AppendReferenceSeparator()
                .Append(name)
                .ToString();
        }

        public string ExternalName((int, int) _, int workbookIndex, string name)
        {
            var sb = new StringBuilder(BOOK_PREFIX_LEN + SHEET_SEPARATOR_LEN + name.Length);
            return sb
                .AppendBookIndex(workbookIndex)
                .AppendReferenceSeparator()
                .Append(name)
                .ToString();
        }

        public string ExternalSheetName((int, int) _, int workbookIndex, string sheet, string name)
        {
            var sb = new StringBuilder(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
            return sb
                .AppendExternalSheetName(workbookIndex, sheet)
                .AppendReferenceSeparator()
                .Append(name)
                .ToString();
        }

        public string BinaryNode((int, int) _, BinaryOperation operation, string leftNode, string rightNode)
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

        public string Unary((int, int) _, UnaryOperation operation, string node)
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

        public string Nested((int, int) _, string node)
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

        protected abstract StringBuilder AppendRefAsText(RowCol cell, (int Row, int Col) point, StringBuilder sb);

        protected abstract StringBuilder AppendRefAsText(ReferenceSymbol reference, (int Row, int Col) point, StringBuilder sb);
    }
}

/// <summary>
/// Extension methods for building formulas.
/// </summary>
internal static class StringBuilderExtensions
{
    public static StringBuilder AppendSheetName(this StringBuilder sb, string sheetName)
    {
        return NameUtils.EscapeName(sb, sheetName);
    }

    public static StringBuilder AppendExternalSheetName(this StringBuilder sb, int workbookIndex, string sheetName)
    {
        if (NameUtils.ShouldQuote(sheetName.AsSpan()))
        {
            return sb
                .Append('\'')
                .AppendBookIndex(workbookIndex)
                .AppendEscapedSheetName(sheetName)
                .Append('\'');
        }

        return sb
            .AppendBookIndex(workbookIndex)
            .AppendSheetName(sheetName);
    }
    public static StringBuilder AppendEscapedSheetName(this StringBuilder sb, string sheetName)
    {
        var startIndex = sb.Length;
        return sb.Append(sheetName).Replace("'", "''", startIndex, sheetName.Length);
    }

    public static StringBuilder AppendReferenceSeparator(this StringBuilder sb)
    {
        return sb.Append('!');
    }

    public static StringBuilder AppendBookIndex(this StringBuilder sb, int bookIndex)
    {
        return sb.Append('[').Append(bookIndex).Append(']');
    }

    public static StringBuilder AppendFunction(this StringBuilder sb, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
    {
        // netstandard 2.0 doesn't have span API for StringBuilder.
        foreach (var c in functionName)
            sb.Append(c);

        return AppendArguments(sb, arguments);
    }

    public static StringBuilder AppendArguments(this StringBuilder sb, IReadOnlyList<string> arguments)
    {
        sb.Append('(');
        if (arguments.Count > 0)
            sb.Append(arguments[0]);

        for (var i = 1; i < arguments.Count; ++i)
            sb.Append(',').Append(arguments[i]);

        return sb.Append(')');
    }
}
