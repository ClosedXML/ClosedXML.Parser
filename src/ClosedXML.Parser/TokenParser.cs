using System;

namespace ClosedXML.Parser;

internal static class TokenParser
{
    private const int MaxCol = 16384;
    private const int MaxRow = 1048576;

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

        startIdx = i; ;
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
}