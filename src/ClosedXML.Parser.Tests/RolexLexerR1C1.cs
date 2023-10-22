using ClosedXML.Parser.Rolex;

namespace ClosedXML.Parser.Tests;

public class RolexLexerR1C1
{
    [Theory]
    [InlineData("r1c1")]
    [InlineData("r[0]c[0]")]
    [InlineData("rc")] // Degenerate R[0]C[0]
    public void Ignores_case_in_references(string formula)
    {
        var tokens = RolexLexer.GetTokensR1C1(formula);
        Assert.Equal(2, tokens.Count);
        Assert.Equal(Token.A1_CELL, tokens[0].SymbolId);
    }

    [Theory]
    [InlineData("R1C1")]
    [InlineData("R[1]C[1]")]
    [InlineData("R1C[1]")]
    [InlineData("R[1]C1")]
    [InlineData("R[1048575]C1")]
    [InlineData("R1C[16383]")]
    [InlineData("R[-1048575]C1")]
    [InlineData("R1C[-16383]")]
    public void Absolute_and_relative_references_can_be_combined(string formula)
    {
        var tokens = RolexLexer.GetTokensR1C1(formula);
        Assert.Equal(2, tokens.Count);
        Assert.Equal(Token.A1_CELL, tokens[0].SymbolId);
    }

    [Theory]
    [InlineData("R[1048576]C1", Token.A1_SPAN_REFERENCE, Token.INTRA_TABLE_REFERENCE, Token.A1_SPAN_REFERENCE, Token.EofSymbolId)]
    [InlineData("R[-1048576]C1", Token.A1_SPAN_REFERENCE, Token.INTRA_TABLE_REFERENCE, Token.A1_SPAN_REFERENCE, Token.EofSymbolId)]
    [InlineData("R1C[16384]", Token.A1_CELL, Token.INTRA_TABLE_REFERENCE, Token.EofSymbolId)]
    [InlineData("R1C[-16384]", Token.A1_CELL, Token.INTRA_TABLE_REFERENCE, Token.EofSymbolId)]
    public void Relative_references_cant_reach_outside_of_worksheet(string formula, params int[] expectedSymbols)
    {
        // Because relative references are one off, i.e. at row 1, the R[1] references second row
        // they can't have full range of columns and rows.
        var tokens = RolexLexer.GetTokensR1C1(formula);
        Assert.Equal(expectedSymbols, tokens.Select(x => x.SymbolId));
    }

    [Fact]
    public void Astral_code_points_are_valid_strings()
    {
        // Grinning face code point U+1F600
        var tokens = RolexLexer.GetTokensR1C1("\"\uD83D\uDE00\"");
        Assert.Equal(new[] {Token.STRING_CONSTANT, Token.EofSymbolId}, tokens.Select(x => x.SymbolId));
    }
}