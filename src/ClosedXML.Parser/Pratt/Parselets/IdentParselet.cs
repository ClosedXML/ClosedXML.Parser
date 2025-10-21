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
        // * name

        // Check for area `A1:B2` or just cell `A1`
        if (_parser.TryLocalAreaA1(token, out var localArea, out var localAreaRange))
        {
            var value = _factory.Reference(ctx, localAreaRange, localArea);
            return new Node<T>(value, localAreaRange);
        }

        // Check for colspan `A:B`
        if (_parser.TryLocalColSpanA1(token, out var localColSpan, out var localColSpanRange))
        {
            var value = _factory.Reference(ctx, localColSpanRange, localColSpan);
            return new Node<T>(value, localColSpanRange);
        }

        // Check for colspan `$1:2`
        if (_parser.TryLocalRowSpanA1(token, out var localRowSpan, out var localRowSpanRange))
        {
            var value = _factory.Reference(ctx, localRowSpanRange, localRowSpan);
            return new Node<T>(value, localRowSpanRange);
        }

        if (_parser.TryGetUnquotedSheet(token, out var sheetNameSpan) && _parser.LookAhead(1).Type == TokenType.Bang)
        {
            // We are now in `sheet!` Parse local reference.
            var sheetName = sheetNameSpan.ToString(); // String allocation, needed for the IAstFactory
            var bangToken = _parser.Consume(TokenType.Bang);
            var sheetWithBangRange = token.Range.ExtendRight(bangToken.Range);

            if (_parser.LookAhead(1) is { Type: TokenType.Ident } sheetRefToken)
            {
                _parser.Consume(TokenType.Ident);

                // Check for area `sheet!A1:B2` or just cell `sheet!A1`
                if (_parser.TryLocalAreaA1(sheetRefToken, out var sheetArea, out var sheetAreaRange))
                {
                    var range = sheetWithBangRange.ExtendRight(sheetAreaRange);
                    var value = _factory.SheetReference(ctx, range, sheetName, sheetArea);
                    return new Node<T>(value, range);
                }

                // Check for colspan `sheet!A:B`
                if (_parser.TryLocalColSpanA1(sheetRefToken, out var sheetColSpan, out var sheetColSpanRange))
                {
                    var range = sheetWithBangRange.ExtendRight(sheetColSpanRange);
                    var value = _factory.SheetReference(ctx, range, sheetName, sheetColSpan);
                    return new Node<T>(value, range);
                }

                // Check for rowspan `sheet!$1:2` The $1 is an ident, but this doesn't detect
                // rowspan starting with a relative row. That is checked below with a token number.
                if (_parser.TryLocalRowSpanA1(sheetRefToken, out var sheetAbsRowSpan, out var sheetAbsRowSpanRange))
                {
                    var range = sheetWithBangRange.ExtendRight(sheetAbsRowSpanRange);
                    var value = _factory.SheetReference(ctx, range, sheetName, sheetAbsRowSpan);
                    return new Node<T>(value, range);
                }

                // Check for colspan `sheet!name`
                if (_parser.TryGetName(sheetRefToken, out var name))
                {
                    var range = sheetWithBangRange.ExtendRight(sheetRefToken.Range);
                    var value = _factory.SheetName(ctx, range, sheetName, name.ToString()); // String allocation, needed for the IAstFactory
                    return new Node<T>(value, range);
                }

                throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
            }

            // Check for rowspan `sheet!1:2` with relative start row
            if (_parser.LookAhead(1).Type == TokenType.Number)
            {
                var sheetRowToken = _parser.Consume(TokenType.Number);
                if (_parser.TryLocalRowSpanA1(sheetRowToken, out var sheetRowSpan, out var sheetRowSpanRange))
                {
                    var range = sheetWithBangRange.ExtendRight(sheetRowSpanRange);
                    var value = _factory.SheetReference(ctx, range, sheetName, sheetRowSpan);
                    return new Node<T>(value, range);
                }
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
}
