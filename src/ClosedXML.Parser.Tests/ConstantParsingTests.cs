using Xunit;

namespace ClosedXML.Parser.Tests;

public class ConstantParsingTests
{
    [Theory]
    [InlineData("1")]
    [InlineData("10.5")]
    [InlineData("10.5E5")]
    [InlineData(".1E-4")]
    public void Number_is_parsed(string formula)
    {
        AssertFormula.CstParsed(formula);
    }

    [Theory]
    [InlineData("#REF!")]
    [InlineData("#VALUE!")]
    public void Error_is_parsed(string formula)
    {
        AssertFormula.CstParsed(formula);
    }
}
