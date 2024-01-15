using static ClosedXML.Parser.ReferenceAxisType;

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
    [Fact]
    public void Parse_row_range()
    {
        // Check A1_ROW ':' A1_ROW path
        AssertAreaReferenceToken("1:1", new ReferenceArea(new RowCol(Relative, 1, None, 0, A1), new RowCol(Relative, 1, None, 0, A1)));
        AssertAreaReferenceToken("$5:10", new ReferenceArea(new RowCol(Absolute, 5, None, 0, A1), new RowCol(Relative, 10, None, 0, A1)));
        AssertAreaReferenceToken("7:$3", new ReferenceArea(new RowCol(Relative, 7, None, 0, A1), new RowCol(Absolute, 3, None, 0, A1)));
        AssertAreaReferenceToken("$1048576:$1048576", new ReferenceArea(new RowCol(Absolute, RowCol.MaxRow, None, 0, A1), new RowCol(Absolute, RowCol.MaxRow, None, 0, A1)));
    }

    [Fact]
    public void Parse_column_range()
    {
        // Check A1_COLUMN ':' A1_COLUMN path
        AssertAreaReferenceToken("A:A", new ReferenceArea(new RowCol(None, 0, Relative, 1, A1), new RowCol(None, 0, Relative, 1, A1)));
        AssertAreaReferenceToken("RW:ST", new ReferenceArea(new RowCol(None, 0, Relative, 491, A1), new RowCol(None, 0, Relative, 514, A1)));
        AssertAreaReferenceToken("$C:D", new ReferenceArea(new RowCol(None, 0, Absolute, 3, A1), new RowCol(None, 0, Relative, 4, A1)));
        AssertAreaReferenceToken("E:$C", new ReferenceArea(new RowCol(None, 0, Relative, 5, A1), new RowCol(None, 0, Absolute, 3, A1)));
        AssertAreaReferenceToken("$XFD:$XFD", new ReferenceArea(new RowCol(None, 0, Absolute, RowCol.MaxCol, A1), new RowCol(None, 0, Absolute, RowCol.MaxCol, A1)));
    }

    private static void AssertAreaReferenceToken(string token, ReferenceArea expectedReference)
    {
        AssertFormula.AssertTokenType(token, FormulaLexer.A1_SPAN_REFERENCE);
        var reference = TokenParser.ParseReference(token, true);
        Assert.Equal(expectedReference, reference);
    }
}