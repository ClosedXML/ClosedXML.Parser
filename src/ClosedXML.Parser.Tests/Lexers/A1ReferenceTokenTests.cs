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
        AssertAreaReferenceToken("$B$3", new ReferenceSymbol(Absolute, 2, Absolute, 3));
        AssertAreaReferenceToken("A1", new ReferenceSymbol(Relative, 1, Relative, 1));
        AssertAreaReferenceToken("XFD1", new ReferenceSymbol(Relative, 16384, Relative, 1));
        AssertAreaReferenceToken("A1048576", new ReferenceSymbol(Relative, 1, Relative, 1048576));
        AssertAreaReferenceToken("$XFD$1048576", new ReferenceSymbol(Absolute, 16384, Absolute, 1048576));
    }

    [Fact]
    public void Parse_row_range()
    {
        // Check A1_ROW ':' A1_ROW path
        AssertAreaReferenceToken("1:1", new ReferenceSymbol(new RowCol(false, 1, true, 1), new RowCol(false, 1, true, MaxCol)));
        AssertAreaReferenceToken("$5:10", new ReferenceSymbol(new RowCol(true, 5, true, 1), new RowCol(false, 10, true, MaxCol)));
        AssertAreaReferenceToken("7:$3", new ReferenceSymbol(new RowCol(false, 7, true, 1), new RowCol(true, 3, true, MaxCol)));
        AssertAreaReferenceToken("$1048576:$1048576", new ReferenceSymbol(new RowCol(true, 1048576, true, 1), new RowCol(true, 1048576, true, MaxCol)));
    }

    [Fact]
    public void Parse_column_range()
    {
        // Check A1_COLUMN ':' A1_COLUMN path
        AssertAreaReferenceToken("A:A", new ReferenceSymbol(new RowCol(true, 1, false, 1), new RowCol(true, MaxRow, false, 1)));
        AssertAreaReferenceToken("RW:ST", new ReferenceSymbol(new RowCol(true, 1, false, 491), new RowCol(true, MaxRow, false, 514)));
        AssertAreaReferenceToken("$C:D", new ReferenceSymbol(new RowCol(true, 1, true, 3), new RowCol(true, MaxRow, false, 4)));
        AssertAreaReferenceToken("E:$C", new ReferenceSymbol(new RowCol(true, 1, false, 5), new RowCol(true, MaxRow, true, 3)));
        AssertAreaReferenceToken("$XFD:$XFD", new ReferenceSymbol(new RowCol(true, 1, true, MaxCol), new RowCol(true, MaxRow, true, MaxCol)));
    }

    [Fact]
    public void Parse_area()
    {
        // Check A1_AREA path
        AssertAreaReferenceToken("A1:A1", new ReferenceSymbol(new RowCol(Relative, 1, Relative, 1)));
        AssertAreaReferenceToken("Z1:AB25", new ReferenceSymbol(new RowCol(false, 1, false, 26), new RowCol(25, 28)));
        AssertAreaReferenceToken("$XFC$1048575:$XFD$1048576", new ReferenceSymbol(new RowCol(true, MaxRow - 1, true, MaxCol - 1), new RowCol(true, MaxRow, true, MaxCol)));
    }

    private static void AssertAreaReferenceToken(string token, ReferenceSymbol expectedReference)
    {
        AssertFormula.AssertTokenType(token, FormulaLexer.A1_REFERENCE);
        var reference = TokenParser.ParseReference(token, true);
        Assert.Equal(expectedReference, reference);
    }
}