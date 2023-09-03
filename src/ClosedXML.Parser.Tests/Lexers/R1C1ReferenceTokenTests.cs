using ClosedXML.Parser.Rolex;

namespace ClosedXML.Parser.Tests.Lexers;

public class R1C1ReferenceTokenTests
{
    [Theory]
    [MemberData(nameof(TestDataOneCorner))]
    [MemberData(nameof(TestDataTwoCorners))]
    public void Parse_extracts_information_from_token(string token, ReferenceSymbol expectedReference)
    {
        Assert.Equal(new[] {Token.A1_REFERENCE, Token.EofSymbolId}, RolexLexer.GetTokensR1C1(token).Select(x => x.SymbolId));
        var reference = TokenParser.ParseReference(token.AsSpan(), false);
        Assert.Equal(expectedReference, reference);
    }

    public static IEnumerable<object[]> TestDataOneCorner
    {
        get
        {
            // The `C` is a shortcut for `C[0]`
            yield return new object[]
            {
                "C",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0))
            };

            yield return new object[]
            {
                "C[-14]",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, -14))
            };

            yield return new object[]
            {
                "C75",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Absolute, 75))
            };

            // The `R` is a shortcut for `R[0]`
            yield return new object[]
            {
                "R",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0))
            };

            yield return new object[]
            {
                "R[-14]",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Relative, -14, ReferenceAxisType.None, 0))
            };

            yield return new object[]
            {
                "R75",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Absolute, 75, ReferenceAxisType.None, 0))
            };

            yield return new object[]
            {
                "RC",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.Relative, 0))
            };

            yield return new object[]
            {
                "R[7]C2",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Relative, 7, ReferenceAxisType.Absolute, 2))
            };

            yield return new object[]
            {
                "R812C[7]",
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Absolute, 812, ReferenceAxisType.Relative, 7))
            };
        }
    }

    public static IEnumerable<object[]> TestDataTwoCorners
    {
        get
        {
            yield return new object[]
            {
                "R1C2:R3C4",
                new ReferenceSymbol(
                    new RowCol(ReferenceAxisType.Absolute, 1, ReferenceAxisType.Absolute, 2),
                    new RowCol(ReferenceAxisType.Absolute, 3, ReferenceAxisType.Absolute, 4))
            };

            yield return new object[]
            {
                "C:R",
                new ReferenceSymbol(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0),
                    new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0))
            };

            yield return new object[]
            {
                "R[-1]C[-2]:R[-3]C[-4]",
                new ReferenceSymbol(
                    new RowCol(ReferenceAxisType.Relative, -1, ReferenceAxisType.Relative, -2),
                    new RowCol(ReferenceAxisType.Relative, -3, ReferenceAxisType.Relative, -4))
            };

            yield return new object[]
            {
                "R:C",
                new ReferenceSymbol(
                    new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0))
            };
        }
    }
}