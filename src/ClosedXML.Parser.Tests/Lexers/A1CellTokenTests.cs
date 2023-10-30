using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests.Lexers;

/// <summary>
/// Test of a parsing of a token <c>A1_CELL</c>.
/// <code>
/// A1_CELL
///     : A1_COLUMN A1_ROW
///     ;
/// </code>
/// </summary>
public class A1CellTokenTests
{
    [Fact]
    public void Parse_a1_cell()
    {
        // Check A1_CELL path
        AssertAreaReferenceToken("$B$3", new ReferenceSymbol(Absolute, 3, Absolute, 2, A1));
        AssertAreaReferenceToken("A1", new ReferenceSymbol(Relative, 1, Relative, 1, A1));
        AssertAreaReferenceToken("XFD1", new ReferenceSymbol(Relative, 1, Relative, 16384, A1));
        AssertAreaReferenceToken("A1048576", new ReferenceSymbol(Relative, 1048576, Relative, 1, A1));
        AssertAreaReferenceToken("$XFD$1048576", new ReferenceSymbol(Absolute, 1048576, Absolute, 16384, A1));
    }

    private static void AssertAreaReferenceToken(string token, ReferenceSymbol expectedReference)
    {
        AssertFormula.AssertTokenType(token, Token.A1_CELL);
        var reference = TokenParser.ParseReference(token, true);
        Assert.Equal(expectedReference, reference);
    }
}