using Xunit;

namespace ClosedXML.Parser.Tests;

public class FunctionTests
{
    [Theory]
    [InlineData("TRUE(TRUE)")]
    public void Ambiguous_built_in_function_name_is_recognized_as_function(string formula)
    {
        AssertFormula.CstParsed(formula);
    }
}