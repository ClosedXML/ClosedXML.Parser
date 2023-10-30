namespace ClosedXML.Parser.Tests.Lexers;

/// <summary>
/// Test of a parsing of a token <c>A1_SPAN_REFERENCE</c>.
/// <code>
/// A1_SPAN_REFERENCE
///        : A1_COLUMN ':' A1_COLUMN
///        | A1_ROW ':' A1_ROW
///        ;
/// </code>
/// </summary>
public class A1SpanReferenceTokenTests
{
    private const int MaxCol = 16384;
    private const int MaxRow = 1048576;

    [Fact]
    public void Parse_row_range()
    {
        // Check A1_ROW ':' A1_ROW path
        AssertAreaReferenceToken("1:1", new ReferenceSymbol(new RowCol(false, 1, true, 1, A1), new RowCol(false, 1, true, MaxCol, A1)));
        AssertAreaReferenceToken("$5:10", new ReferenceSymbol(new RowCol(true, 5, true, 1, A1), new RowCol(false, 10, true, MaxCol, A1)));
        AssertAreaReferenceToken("7:$3", new ReferenceSymbol(new RowCol(false, 7, true, 1, A1), new RowCol(true, 3, true, MaxCol, A1)));
        AssertAreaReferenceToken("$1048576:$1048576", new ReferenceSymbol(new RowCol(true, 1048576, true, 1, A1), new RowCol(true, 1048576, true, MaxCol, A1)));
    }

    [Fact]
    public void Parse_column_range()
    {
        // Check A1_COLUMN ':' A1_COLUMN path
        AssertAreaReferenceToken("A:A", new ReferenceSymbol(new RowCol(true, 1, false, 1, A1), new RowCol(true, MaxRow, false, 1, A1)));
        AssertAreaReferenceToken("RW:ST", new ReferenceSymbol(new RowCol(true, 1, false, 491, A1), new RowCol(true, MaxRow, false, 514, A1)));
        AssertAreaReferenceToken("$C:D", new ReferenceSymbol(new RowCol(true, 1, true, 3, A1), new RowCol(true, MaxRow, false, 4, A1)));
        AssertAreaReferenceToken("E:$C", new ReferenceSymbol(new RowCol(true, 1, false, 5, A1), new RowCol(true, MaxRow, true, 3, A1)));
        AssertAreaReferenceToken("$XFD:$XFD", new ReferenceSymbol(new RowCol(true, 1, true, MaxCol, A1), new RowCol(true, MaxRow, true, MaxCol, A1)));
    }
    
    private static void AssertAreaReferenceToken(string token, ReferenceSymbol expectedReference)
    {
        AssertFormula.AssertTokenType(token, FormulaLexer.A1_SPAN_REFERENCE);
        var reference = TokenParser.ParseReference(token, true);
        Assert.Equal(expectedReference, reference);
    }
}