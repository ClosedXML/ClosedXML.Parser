using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests.Lexers;

/// <summary>
/// Test of a parsing of a token <c>A1_REFERENCE</c>.
/// <code>
/// A1_REFERENCE
///        : A1_COLUMN ':' A1_COLUMN
///        | A1_ROW ':' A1_ROW
///        | A1_CELL
///        | A1_AREA
///        ;
/// </code>
/// </summary>
public class A1ReferenceTokenTests
{
    private const int MaxCol = 16384;
    private const int MaxRow = 1048576;

    [Fact]
    public void Parse_a1_cell()
    {
        // Check A1_CELL path
        AssertAreaReferenceToken("$B$3", new ReferenceArea(Absolute, 2, Absolute, 3));
        AssertAreaReferenceToken("A1", new ReferenceArea(Relative, 1, Relative, 1));
        AssertAreaReferenceToken("XFD1", new ReferenceArea(Relative, 16384, Relative, 1));
        AssertAreaReferenceToken("A1048576", new ReferenceArea(Relative, 1, Relative, 1048576));
        AssertAreaReferenceToken("$XFD$1048576", new ReferenceArea(Absolute, 16384, Absolute, 1048576));
    }

    [Fact]
    public void Parse_row_range()
    {
        // Check A1_ROW ':' A1_ROW path
        AssertAreaReferenceToken("1:1", new ReferenceArea(new RowCol(true, 1, false, 1), new RowCol(true, MaxCol, false, 1)));
        AssertAreaReferenceToken("$5:10", new ReferenceArea(new RowCol(true, 1, true, 5), new RowCol(true, MaxCol, false, 10)));
        AssertAreaReferenceToken("7:$3", new ReferenceArea(new RowCol(true, 1, false, 7), new RowCol(true, MaxCol, true, 3)));
        AssertAreaReferenceToken("$1048576:$1048576", new ReferenceArea(new RowCol(true, 1, true, 1048576), new RowCol(true, MaxCol, true, 1048576)));
    }

    [Fact]
    public void Parse_column_range()
    {
        // Check A1_COLUMN ':' A1_COLUMN path
        AssertAreaReferenceToken("A:A", new ReferenceArea(new RowCol(false, 1, true, 1), new RowCol(false, 1, true, MaxRow)));
        AssertAreaReferenceToken("RW:ST", new ReferenceArea(new RowCol(false, 491, true, 1), new RowCol(false, 514, true, MaxRow)));
        AssertAreaReferenceToken("$C:D", new ReferenceArea(new RowCol(true, 3, true, 1), new RowCol(false, 4, true, MaxRow)));
        AssertAreaReferenceToken("E:$C", new ReferenceArea(new RowCol(false, 5, true, 1), new RowCol(true, 3, true, MaxRow)));
        AssertAreaReferenceToken("$XFD:$XFD", new ReferenceArea(new RowCol(true, MaxCol, true, 1), new RowCol(true, MaxCol, true, MaxRow)));
    }

    [Fact]
    public void Parse_area()
    {
        // Check A1_AREA path
        AssertAreaReferenceToken("A1:A1", new ReferenceArea(new RowCol(Relative, 1, Relative, 1)));
        AssertAreaReferenceToken("Z1:AB25", new ReferenceArea(new RowCol(false, 26, false, 1), new RowCol(28, 25)));
        AssertAreaReferenceToken("$XFC$1048575:$XFD$1048576", new ReferenceArea(new RowCol(true, MaxCol - 1, true, MaxRow - 1), new RowCol(true, MaxCol, true, MaxRow)));
    }

    private static void AssertAreaReferenceToken(string token, ReferenceArea expectedReference)
    {
        AssertFormula.AssertTokenType(token, FormulaLexer.A1_REFERENCE);
        var reference = TokenParser.ParseReference(token, true);
        Assert.Equal(expectedReference, reference);
    }
}