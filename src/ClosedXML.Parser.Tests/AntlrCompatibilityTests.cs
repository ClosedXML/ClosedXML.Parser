using ClosedXML.Parser.Rolex;

namespace ClosedXML.Parser.Tests;

/// <summary>
/// ANTLR parser is the source of truth. This test class checks that ANTLR output and Rolex/RDP have same output.
/// </summary>
public class AntlrCompatibilityTests
{
    [Theory]
    [InlineData("./data/enron/formulas.csv")]
    [InlineData("./data/euses/formulas.csv")]
    [InlineData("./data/contributions/formulas.csv")]
    public void Produce_same_tokens_for_data_sets(string dataSetFile)
    {
        foreach (var formula in DataSets.ReadCsv(dataSetFile))
        {
            var antlrTokens = AssertFormula.GetAntlrTokens(formula);
            var rolexTokens = RolexLexer.GetTokensA1(formula.AsSpan());

            Assert.Equal(antlrTokens, rolexTokens);
        }
    }
}


