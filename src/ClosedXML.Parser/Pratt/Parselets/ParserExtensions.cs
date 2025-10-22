using System;
using static ClosedXML.Parser.Pratt.CompatUtils;

namespace ClosedXML.Parser.Pratt.Parselets;

internal static class ParserExtensions
{
    private const int MIN_A1_LENGTH = 2; // A1
    private const int MAX_A1_LENGTH = 1 + 3 + 1 + 7; // $XFD$1048576
    private const int MIN_COL_LENGTH = 1; // A
    private const int MAX_COL_LENGTH = 4; // $XFD
    private const int MIN_ROW_LENGTH = 1; // 1
    private const int MAX_ROW_LENGTH = 8; // $1048576

    public static bool TryReferenceA1<T, TContext>(this Parser<T, TContext> parser, Token token, out ReferenceArea area, out SymbolRange range)
    {
        if (token.Type is not TokenType.Ident and not TokenType.Number)
        {
            area = default;
            range = default;
            return false;
        }

        // Check for area `A1:B2` or just cell `A1`
        if (parser.TryLocalAreaA1(token, out area, out range))
            return true;

        // Check for colspan `A:B`
        if (parser.TryLocalColSpanA1(token, out area, out range))
            return true;

        // Check for rowspan `1:2`, can be ident or number token
        if (parser.TryLocalRowSpanA1(token, out area, out range))
            return true;

        return false;
    }

    public static bool TryLocalAreaA1<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReferenceArea area, out SymbolRange range)
    {
        if (identToken.Type != TokenType.Ident)
        {
            area = default;
            range = default;
            return false;
        }

        var ident = identToken.GetText(parser.Input);

        if (TryGetCellA1(ident, out var cell1))
        {
            if (parser.LookAhead(1).Type == TokenType.Range &&
                parser.LookAhead(2) is { Type: TokenType.Ident } maybeCell2Token &&
                TryGetCellA1(maybeCell2Token.GetText(parser.Input), out var cell2))
            {
                // Result: area A1:B2
                // The code is joining two cells into an area through range operator, but that
                // is allowed. Range is highest priority operator, left to right associativity.
                var rangeToken = parser.Consume(TokenType.Range);
                var cell2Token = parser.Consume(TokenType.Ident);

                area = new ReferenceArea(cell1, cell2);
                range = identToken.Range
                    .ExtendRight(rangeToken.Range)
                    .ExtendRight(cell2Token.Range);
                return true;
            }

            // Result: cell A1
            area = new ReferenceArea(cell1);
            range = identToken.Range;
            return true;
        }

        range = default;
        area = default;
        return false;
    }

    public static bool TryLocalColSpanA1<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReferenceArea area, out SymbolRange range)
    {
        if (identToken.Type != TokenType.Ident)
        {
            area = default;
            range = default;
            return false;
        }

        var ident = identToken.GetText(parser.Input);

        // Careful, 'A' can be just a name without the other column
        if (TryGetColA1(ident, out var col1) &&
            parser.LookAhead(1).Type == TokenType.Range &&
            parser.LookAhead(2) is { Type: TokenType.Ident } maybeCol2Token &&
            TryGetColA1(maybeCol2Token.GetText(parser.Input), out var col2))
        {
            // Result: colspan A:B
            var rangeToken = parser.Consume(TokenType.Range);
            var col2Token = parser.Consume(TokenType.Ident);

            area = new ReferenceArea(col1, col2);
            range = identToken.Range
                .ExtendRight(rangeToken.Range)
                .ExtendRight(col2Token.Range);
            return true;
        }

        area = default;
        range = default;
        return false;
    }

    public static bool TryLocalRowSpanA1<T, TContext>(this Parser<T, TContext> parser, Token numberOrIdentToken, out ReferenceArea area, out SymbolRange range)
    {
        if (numberOrIdentToken.Type is not TokenType.Ident and not TokenType.Number)
        {
            area = default;
            range = default;
            return false;
        }

        var numberOrIdent = numberOrIdentToken.GetText(parser.Input);

        if (TryGetRowA1(numberOrIdent, out var row1) &&
            parser.LookAhead(1).Type == TokenType.Range &&
            parser.LookAhead(2) is { Type: TokenType.Number or TokenType.Ident } maybeRow2Token &&
            TryGetRowA1(maybeRow2Token.GetText(parser.Input), out var row2))
        {
            // Result: rowspan 1:2
            var rangeToken = parser.Consume(TokenType.Range);
            var row2Token = parser.Consume();

            area = new ReferenceArea(row1, row2);
            range = numberOrIdentToken.Range
                .ExtendRight(rangeToken.Range)
                .ExtendRight(row2Token.Range);
            return true;
        }

        area = default;
        range = default;
        return false;
    }

    public static bool TryGetUnquotedSheet<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReadOnlySpan<char> sheetName)
    {
        var text = identToken.GetText(parser.Input);
        var isUnquotedSheet = NameUtils.IsSheetNameValid(text) && !NameUtils.ShouldQuote(text);
        if (isUnquotedSheet)
        {
            sheetName = text;
            return true;
        }

        sheetName = default;
        return false;
    }

    public static bool TryGetName<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReadOnlySpan<char> name)
    {
        if (identToken.Type != TokenType.Ident)
        {
            name = default;
            return false;
        }

        var text = identToken.GetText(parser.Input);
        if (NameUtils.IsNameValid(text))
        {
            name = text;
            return true;
        }

        name = default;
        return false;
    }

    /// <summary>
    /// Is the <paramref name="text"/> a valid A1 cell reference? No padding, case insensitive.
    /// </summary>
    public static bool TryGetCellA1(ReadOnlySpan<char> text, out RowCol cell)
    {
        cell = default;
        if (text.Length is < MIN_A1_LENGTH or > MAX_A1_LENGTH)
            return false;

        var i = 0;
        var absCol = text[i] == '$';
        if (absCol) ++i;

        var col = 0;
        while (i < text.Length && IsAsciiLetter(text[i]))
            col = col * 26 + GetColIndex(text[i++]) + 1;

        if (col is < RowCol.MinCol or > RowCol.MaxCol || i >= text.Length)
            return false;

        var absRow = text[i] == '$';
        if (absRow)
        {
            if (++i >= text.Length)
                return false;
        }

        if (text[i] == '0')
            return false;

        var row = 0;
        while (i < text.Length && IsAsciiDigit(text[i]))
            row = row * 10 + text[i++] - '0';

        if (row is < RowCol.MinRow or > RowCol.MaxRow || i < text.Length)
            return false;

        cell = new RowCol(
            absRow ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, row,
            absCol ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, col,
            A1);
        return true;
    }

    /// <summary>
    /// Is the <paramref name="text"/> a valid end of an A1 colspan? No padding, case insensitive.
    /// Valid examples: <c>A</c>, <c>a</c>, <c>$A</c>, <c>$XFD</c>.
    /// Invalid examples: <c> A </c>, <c>$ a</c>, <c>$</c>, <c>$XFE</c>.
    /// </summary>
    public static bool TryGetColA1(ReadOnlySpan<char> text, out RowCol colRef)
    {
        colRef = default;
        if (text.Length is < MIN_COL_LENGTH or > MAX_COL_LENGTH)
            return false;

        var i = 0;
        var absCol = text[i] == '$';
        if (absCol) ++i;

        var col = 0;
        while (i < text.Length && IsAsciiLetter(text[i]))
            col = col * 26 + GetColIndex(text[i++]) + 1;

        if (col is < RowCol.MinCol or > RowCol.MaxCol || i < text.Length)
            return false;

        colRef = new RowCol(
            ReferenceAxisType.None, 0,
            absCol ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, col,
            A1);
        return true;
    }

    /// <summary>
    /// Is the <paramref name="text"/> a valid end of an A1 rowspan? No padding.
    /// Valid examples: <c>1</c>, <c>$1</c>, <c>$1048576</c>.
    /// Invalid examples: <c>1.0</c>, <c>$ 1</c>, <c>$</c>, <c>$1048577</c>.
    /// </summary>
    public static bool TryGetRowA1(ReadOnlySpan<char> text, out RowCol rowRef)
    {
        rowRef = default;
        if (text.Length is < MIN_ROW_LENGTH or > MAX_ROW_LENGTH)
            return false;

        var i = 0;
        var absRow = text[i] == '$';
        if (absRow)
        {
            if (++i >= text.Length)
                return false;
        }

        if (text[i] == '0')
            return false;

        var row = 0;
        while (i < text.Length && IsAsciiDigit(text[i]))
            row = row * 10 + text[i++] - '0';

        if (row is < RowCol.MinRow or > RowCol.MaxRow || i < text.Length)
            return false;

        rowRef = new RowCol(
            absRow ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, row,
            ReferenceAxisType.None, 0,
            A1);
        return true;
    }

    private static int GetColIndex(char asciiLetter)
    {
        return (asciiLetter | 0x20) - 'a';
    }
}
