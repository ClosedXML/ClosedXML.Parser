using System;
using System.Diagnostics;

namespace ClosedXML.Parser;

internal static class TokenParser
{
    private const int MaxCol = 16384;
    private const int MaxRow = 1048576;

    /// <summary>
    /// Parse <see cref="FormulaLexer.SINGLE_SHEET_PREFIX"/> token.
    /// </summary>
    internal static void ParseSingleSheetPrefix(ReadOnlySpan<char> input, out int? index, out string sheetName)
    {
        var isEscaped = input[0] == '\'';
        input = isEscaped
            ? input.Slice(1, input.Length - 3) // second sheet name ends with TICK EXCLAMATION_MARK ('!). 
            : input.Slice(0, input.Length - 1); // only strip exclamation mark

        // Parse optional WORKBOOK_INDEX
        input = ExtractWorkbookIndex(input, out index);

        var sheetRangeSpan = input;
        if (!isEscaped)
        {
            // SHEET_NAME
            sheetName = sheetRangeSpan.ToString();
            return;
        }

        // The ending '! have been stripped from escape
        sheetName = GetEscapedSheetName(input);
    }

    /// <summary>
    /// Parse token <see cref="FormulaLexer.SHEET_RANGE_PREFIX"/>
    /// </summary>
    internal static void ParseSheetRangePrefix(ReadOnlySpan<char> input, out int? index, out string firstSheetName, out string secondSheetName)
    {
        var isEscaped = input[0] == '\'';
        input = isEscaped
            ? input.Slice(1, input.Length - 3) // second sheet name ends with TICK EXCLAMATION_MARK ('!). 
            : input.Slice(0, input.Length - 1); // only strip exclamation mark

        // Parse optional WORKBOOK_INDEX
        input = ExtractWorkbookIndex(input, out index);

        var sheetRangeSpan = input;
        if (!isEscaped)
        {
            // SHEET_NAME ':' SHEET_NAME
            var endIndex = sheetRangeSpan.IndexOf(':');
            firstSheetName = sheetRangeSpan.Slice(0, endIndex).ToString();
            secondSheetName = sheetRangeSpan.Slice(endIndex + 1).ToString();
            return;
        }

        // Parse SHEET_NAME_SPECIAL which can contain escaped tick (') as double tick
        firstSheetName = GetEscapedSheetName(ref input, ':'); // Even escaped sheet name can't contain :
        secondSheetName = GetEscapedSheetName(input);
    }

    private static ReadOnlySpan<char> ExtractWorkbookIndex(ReadOnlySpan<char> input, out int? wbIndex)
    {
        if (input[0] != '[')
        {
            wbIndex = null;
            return input;
        }

        var i = 0;
        var number = 0;
        var c = input[++i];
        do
        {
            number = number * 10 + c - '0';
            c = input[++i];
        } while (c != ']');

        wbIndex = number;
        return input.Slice(i + 1);
    }

    private static string GetEscapedSheetName(ref ReadOnlySpan<char> input, char endChar)
    {
        Span<char> buffer = stackalloc char[input.Length];
        var bufferIdx = 0;
        var inputIdx = 0;
        do
        {
            if (input[inputIdx] == '\'')
                inputIdx++;

            buffer[bufferIdx++] = input[inputIdx++];
        } while (input[inputIdx] != endChar);

        input = input.Slice(inputIdx + 1);
        return buffer.Slice(0, bufferIdx).ToString();
    }

    private static string GetEscapedSheetName(ReadOnlySpan<char> input)
    {
        Span<char> buffer = stackalloc char[input.Length];
        var bufferIdx = 0;
        var inputIdx = 0;
        do
        {
            if (input[inputIdx] == '\'')
                inputIdx++;

            buffer[bufferIdx++] = input[inputIdx++];
        } while (input.Length > inputIdx);

        return buffer.Slice(0, bufferIdx).ToString();
    }

    internal static ReadOnlySpan<char> ExtractLocalFunctionName(ReadOnlySpan<char> functionNameWithBrace)
    {
        // In most cases, there won't be any whitespace
        var endPosition = functionNameWithBrace[functionNameWithBrace.Length - 1] == '('
            ? functionNameWithBrace.Length - 1
            : functionNameWithBrace.LastIndexOf('(');
        var functionName = functionNameWithBrace.Slice(0, endPosition);
        return functionName;
    }

    /// <summary>
    /// Extract info about cell reference from a <c>A1_REFERENCE</c> token.
    /// </summary>
    internal static CellArea ParseA1Reference(ReadOnlySpan<char> input)
    {
        // The point of this method is to be fast, not pretty. It assumes that input has
        // already been checked by lexer and thus will never be incorrect.
        var i = 0;
        var abs1 = IsAbsolute(input, i);
        if (abs1)
            i++;

        var colStart = input[i] >= 'A' && input[i] <= 'Z';
        if (!colStart)
        {
            // A1_ROW ':' A1_ROW
            var row1 = ReadRow(input, ref i);
            i++; // Skip ':'
            var absRow2 = IsAbsolute(input, i);
            if (absRow2)
                i++; // Skip '$'

            var row2 = ReadRow(input, ref i);
            return new CellArea(new CellReference(true, 1, abs1, row1), new CellReference(true, MaxCol, absRow2, row2));
        }

        var col = ReadColumn(input, ref i);
        if (input[i] == ':')
        {
            // A1_COLUMN ':' A1_COLUMN
            i++; // Skip ':'
            var absCol2 = IsAbsolute(input, i);
            if (absCol2)
                i++;

            var col2 = ReadColumn(input, ref i);
            return new CellArea(new CellReference(abs1, col, true, 1),
                new CellReference(absCol2, col2, true, MaxRow));
        }

        var secondAbsolute = false;
        if (IsAbsolute(input, i))
        {
            secondAbsolute = true;
            i++;
        }

        // A1_CELL | A1_AREA : A1_CELL ':' A1_CELL
        var row = ReadRow(input, ref i);

        var cell = new CellReference(abs1, col, secondAbsolute, row);
        if (i == input.Length)
        {
            // A1_CELL
            return new CellArea(cell, cell);
        }

        // A1_AREA
        i++; // Skip ':'
        var secondCell = ReadCell(input, ref i);
        return new CellArea(cell, secondCell);
    }

    private static CellReference ReadCell(ReadOnlySpan<char> input, ref int i)
    {
        var colAbs = IsAbsolute(input, i);
        if (colAbs)
            i++;

        var col = ReadColumn(input, ref i);
        var rowAbs = IsAbsolute(input, i);
        if (rowAbs)
            i++;

        var row = ReadRow(input, ref i);
        return new CellReference(colAbs, col, rowAbs, row);
    }

    private static bool IsAbsolute(ReadOnlySpan<char> input, int startIdx) => input[startIdx] == '$';

    // Call only when first char is column
    private static int ReadColumn(ReadOnlySpan<char> input, ref int startIdx)
    {
        var column = 0;
        var i = startIdx;

        do
        {
            var letter = input[i] - 'A' + 1;
            column = column * 26 + letter;
            i++;
        } while (i < input.Length && input[i] >= 'A' && input[i] <= 'Z');

        startIdx = i;
        return column;
    }

    private static int ReadRow(ReadOnlySpan<char> input, ref int startIdx)
    {
        var row = 0;
        var i = startIdx;
        do
        {
            var digit = input[i] - '0';
            row = row * 10 + digit;
            i++;
        } while (i < input.Length && input[i] >= '0' && input[i] <= '9');

        startIdx = i;
        return row;
    }

    internal static void ParseIntraTableReference(ReadOnlySpan<char> input, out StructuredReferenceArea area, out string? firstColumn, out string? lastColumn)
    {
        // Skip first char, it's always '['
        var i = 1;
        if (input[i] == '#')
        {
            // Pattern is a KEYWORD
            area = GetArea(input, i);
            firstColumn = null;
            lastColumn = null;
            return;
        }

        if (input[i] != '[' && input[i] != ' ')
        {
            // Pattern is '[]', '[First]' or '[First:Last]' or '[First:[Last]]'
            // because simple column can't start with a space.
            area = StructuredReferenceArea.None;
            if (input[i] == ']')
            {
                // Pattern is '[]', i.e. whole table.
                firstColumn = null;
                lastColumn = null;
                return;
            }

            // Read simple column
            i = GetStructuredName(input, i, out firstColumn);
            if (i < input.Length && input[i] == ':')
                GetStructuredName(input, i + 1, out lastColumn);
            else
                lastColumn = null;

            return;
        }

        // Pattern is SPACED_LBRACKET INNER_REFERENCE SPACED_RBRACKET

        // Skip potential whitespaces at the beginning of a structured reference (SPACED_LBRACKET)
        i = SkipWhitespaces(input, i);
        area = StructuredReferenceArea.None;
        if (input[i + 1] == '#')
        {
            // Inner reference contains a keyword.
            var listItem = GetArea(input, ++i);
            i += GetLength(listItem) + 1;
            area |= listItem;

            i = SkipComma(input, i);
        }

        if (input[i + 1] == '#')
        {
            // Item is a keyword list, either
            // * '[#Headers]' SPACED_COMMA '[#Data]'
            // * '[#Data]' SPACED_COMMA '[#Totals]'
            var listItem = GetArea(input, ++i);
            i += GetLength(listItem) + 1;
            area |= listItem;

            i = SkipComma(input, i);
        }

        // KEYWORD_LIST can contain at most two item specifiers.
        // After keyword list, we get either a COLUMN or a COLUMN:COLUMN
        i = GetStructuredName(input, i, out firstColumn);
        if (i < input.Length && input[i] == ':')
            GetStructuredName(input, i + 1, out lastColumn);
        else
            lastColumn = null;
    }

    private static int SkipComma(ReadOnlySpan<char> input, int i)
    {
        // comma might be wrapped in whitespaces.
        i = SkipWhitespaces(input, i);

        Debug.Assert(input[i] == ',');
        i++;
        i = SkipWhitespaces(input, i);
        return i;
    }

    private static int SkipWhitespaces(ReadOnlySpan<char> input, int i)
    {
        for (; i < input.Length; i++)
        {
            if (!IsWhiteSpace(input[i]))
                break;
        }

        return i;
    }

    /// <summary>
    /// Read a structured name until the end bracket or column
    /// </summary>
    /// <param name="input">Input span.</param>
    /// <param name="startIdx">First index of expected name. It will either contain a bracket or first letter of column name.</param>
    /// <param name="columnName">Parsed name.</param>
    private static int GetStructuredName(ReadOnlySpan<char> input, int startIdx, out string columnName)
    {
        Span<char> buffer = stackalloc char[input.Length];
        var bufferIdx = 0;
        var i = startIdx + (input[startIdx] == '[' ? 1 : 0);
        var c = input[i];
        for (; c is not ']' and not ':'; c = input[++i])
        {
            if (c == '\'')
                c = input[++i];

            buffer[bufferIdx++] = c;
        }

        columnName = buffer.Slice(0, bufferIdx).ToString();
        return i + (c == ']' ? 1 : 0); // char after last bracket
    }

    private static StructuredReferenceArea GetArea(ReadOnlySpan<char> input, int i)
    {
        var item = input[i + 1] switch
        {
            'A' => StructuredReferenceArea.All,
            'a' => StructuredReferenceArea.All,
            'D' => StructuredReferenceArea.Data,
            'd' => StructuredReferenceArea.Data,
            'H' => StructuredReferenceArea.Headers,
            'h' => StructuredReferenceArea.Headers,
            'T' => input[i + 2] switch
            {
                'O' => StructuredReferenceArea.Totals,
                'o' => StructuredReferenceArea.Totals,
                'H' => StructuredReferenceArea.ThisRow,
                'h' => StructuredReferenceArea.ThisRow,
                _ => throw new NotSupportedException()
            },
            _ => throw new NotSupportedException()
        };
        return item;
    }

    private static int GetLength(StructuredReferenceArea item)
    {
        return item switch
        {
            StructuredReferenceArea.All => 4,
            StructuredReferenceArea.Data => 5,
            StructuredReferenceArea.Headers => 8,
            StructuredReferenceArea.ThisRow => 9,
            StructuredReferenceArea.Totals => 7,
            _ => throw new InvalidOperationException()
        };
    }

    private static bool IsWhiteSpace(char c)
    {
        return c is ' ' or '\n' or '\r';
    }

    internal static int ParseBookPrefix(ReadOnlySpan<char> input)
    {
        ExtractWorkbookIndex(input, out var index);
        Debug.Assert(index.HasValue);
        return index.Value;
    }
}