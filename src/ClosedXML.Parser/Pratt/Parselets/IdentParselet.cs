using System;

namespace ClosedXML.Parser.Pratt.Parselets;

internal class IdentParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public IdentParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        _factory = factory;
        _parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        // When we receive an ident, there are following possibilities what it could be (checked
        // in this order):
        // * A1:B2
        // * A1
        // * A:B
        // * $4:6 - rowspan starting with an absolute row
        // * sheet!A1:A2
        // * sheet!A1
        // * sheet!A:B
        // * sheet!$1:2
        // * sheet!name
        // * sheet!1:2
        // * TRUE/FALSE
        // * sheet1:sheet2!A1:A2
        // * sheet1:sheet2!A1
        // * sheet1:sheet2!A:B
        // * sheet1:sheet2!$1:2
        // * name

        // Check for area `A1:B2` or just cell `A1`
        // Check for colspan `A:B`
        // Check for colspan `$1:2` with absolute row start, because this is an "ident" prefix parselet
        if (_parser.TryReferenceA1(token, out var localArea, out var localAreaRange))
        {
            var value = _factory.Reference(ctx, localAreaRange, localArea);
            return new Node<T>(value, localAreaRange);
        }


        if (_parser.TryGetUnquotedSheet(token, out var sheetNameSpan) && _parser.LookAhead(1).Type == TokenType.Bang)
        {
            // We are now in `sheet!` Parse local reference.
            var sheetName = sheetNameSpan.ToString(); // String allocation, needed for the IAstFactory
            var bangToken = _parser.Consume(TokenType.Bang);
            var sheetWithBangRange = token.Range.ExtendRight(bangToken.Range);

            // No need to check for token type, if EoF, nothing will be matched to such token
            var sheetRefToken = _parser.Consume();
            
            // Check for area `sheet!A1:B2` or just cell `sheet!A1`
            // Check for colspan `sheet!A:B`
            // Check for rowspan `sheet!1:2` with absolute or relative start row
            if (_parser.TryReferenceA1(sheetRefToken, out var sheetArea, out var sheetAreaRange))
            {
                var range = sheetWithBangRange.ExtendRight(sheetAreaRange);
                var value = _factory.SheetReference(ctx, range, sheetName, sheetArea);
                return new Node<T>(value, range);
            }

            // Check for `sheet!name`
            if (_parser.TryGetName(sheetRefToken, out var name))
            {
                var range = sheetWithBangRange.ExtendRight(sheetRefToken.Range);
                var value = _factory.SheetName(ctx, range, sheetName, name.ToString()); // String allocation, needed for the IAstFactory
                return new Node<T>(value, range);
            }

            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        var tokenText = token.GetText(_parser.Input);
        if (EqualCaseInsensitive(tokenText, "TRUE"))
        {
            var value = _factory.LogicalNode(ctx, token.Range, true);
            return new Node<T>(value, token.Range);
        }

        if (EqualCaseInsensitive(tokenText, "FALSE"))
        {
            var value = _factory.LogicalNode(ctx, token.Range, false);
            return new Node<T>(value, token.Range);
        }

        // Check for 3D reference for unquoted sheets:
        // * Sheet1:Sheet2!A1:B2
        // * Sheet1:Sheet2!A1
        // * Sheet1:Sheet2!A:B
        // * Sheet1:Sheet2!1:2
        if (_parser.TryGetUnquotedSheet(token, out var startSheet) &&
            _parser.LookAhead(1).Type == TokenType.Range &&
            _parser.LookAhead(2) is { Type: TokenType.Ident } maybeEndSheetToken &&
            _parser.TryGetUnquotedSheet(maybeEndSheetToken, out var endSheet) &&
            _parser.LookAhead(3).Type == TokenType.Bang)
        {
            var sheetStartToken = token;
            var rangeToken = _parser.Consume(TokenType.Range);
            var sheetEndToken = _parser.Consume(TokenType.Ident);
            var bangToken = _parser.Consume(TokenType.Bang);
            var refToken = _parser.Consume();

            if (_parser.TryReferenceA1(refToken, out var sheetRangeReference, out var sheetRangeReferenceRange))
            {
                var range = sheetStartToken.Range
                    .ExtendRight(rangeToken.Range)
                    .ExtendRight(sheetEndToken.Range)
                    .ExtendRight(bangToken.Range)
                    .ExtendRight(sheetRangeReferenceRange);
                var startSheetString = startSheet.ToString(); // String allocation for the IAstFactory
                var endSheetString = endSheet.ToString();
                var value = _factory.Reference3D(ctx, range, startSheetString, endSheetString, sheetRangeReference);
                return new Node<T>(value, range);
            }

            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        // Check for rowspan `name`
        if (_parser.TryGetName(token, out var workbookName))
        {
            var value = _factory.Name(ctx, token.Range, workbookName.ToString()); // String allocation, needed for the IAstFactory
            return new Node<T>(value, token.Range);
        }

        throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
    }

    private static bool EqualCaseInsensitive(ReadOnlySpan<char> text, string other)
    {
        if (text.Length != other.Length)
            return false;

        return text.CompareTo(other.AsSpan(), StringComparison.OrdinalIgnoreCase) == 0;
    }
}
