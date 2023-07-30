namespace ClosedXML.Parser.Tests.Rules;

public class ConstantRuleTests
{
    [Theory]
    [InlineData("#DIV/0!")]
    [InlineData("#N/A")]
    [InlineData("#NAME?")]
    [InlineData("#NULL!")]
    [InlineData("#NUM!")]
    [InlineData("#VALUE!")]
    [InlineData("#GETTING_DATA")]
    public void NonRefErrors(string error)
    {
        AssertFormula.SingleNodeParsed(error, new ValueNode("Error", error));
        AssertFormula.SingleNodeParsed(error.ToLowerInvariant(), new ValueNode("Error", error));
    }

    [Theory]
    [InlineData("TRUE", true)]
    [InlineData("FALSE", false)]
    public void LogicalConstant(string formula, bool value)
    {
        AssertFormula.SingleNodeParsed(formula, new ValueNode("Logical", value));
        AssertFormula.SingleNodeParsed(formula.ToLowerInvariant(), new ValueNode("Logical", value));
    }

    [Theory]
    [InlineData("1.5e2", 150.0)]
    [InlineData("25.0e-2", 0.25)]
    [InlineData("1", 1.0)]
    [InlineData("5.4", 5.4)]
    public void NumericalConstant(string formula, double value)
    {
        AssertFormula.SingleNodeParsed(formula, new ValueNode("Number", value));
        AssertFormula.SingleNodeParsed(formula.ToUpperInvariant(), new ValueNode("Number", value));
    }

    [Theory]
    [InlineData("\"Hello\"", "Hello")]
    [InlineData("\"Tom \"\"Ben\"\"\"", "Tom \"Ben\"")]
    [InlineData("\"\"", "")]
    public void StringConstant(string formula, string text)
    {
        AssertFormula.SingleNodeParsed(formula, new ValueNode("Text", text));
    }

    [Fact]
    public void Single_element_array()
    {
        AssertFormula.SingleNodeParsed("{1}", new ArrayNode(1, 1, new[] { new ScalarValue(1) }));
    }

    [Fact]
    public void Array_can_contain_number_logical_text_or_error()
    {
        AssertFormula.SingleNodeParsed("{ 1.5 , true , \"Test\" , #n/a }", new ArrayNode(1, 4, new[]
        {
            new ScalarValue(1.5),
            new ScalarValue(true),
            new ScalarValue("Test"),
            new ScalarValue("Error", "#N/A"),
        }));
    }

    [Fact]
    public void Number_in_array_can_have_plus_prefix()
    {
        AssertFormula.SingleNodeParsed("{+3}", new ArrayNode(1, 1, new[] { new ScalarValue(3) }));
    }

    [Theory]
    [InlineData("#REF!")]
    [InlineData("#DIV/0!")]
    [InlineData("#N/A")]
    [InlineData("#NAME?")]
    [InlineData("#NULL!")]
    [InlineData("#NUM!")]
    [InlineData("#VALUE!")]
    [InlineData("#GETTING_DATA")]
    public void Array_can_contain_errors(string error)
    {
        AssertFormula.SingleNodeParsed($"{{{error}}}", new ArrayNode(1, 1, new[] { new ScalarValue("Error", error) }));
    }

    [Fact]
    public void Array_cant_contain_blanks()
    {
        AssertFormula.CheckParsingErrorContains("{1,,}", " Unexpected token COMMA.");
    }

    [Fact]
    public void Empty_array_is_unparsable()
    {
        AssertFormula.CheckParsingErrorContains("{}", "Unexpected token CLOSE_CURLY.");
    }

    [Fact]
    public void Rows_of_array_must_have_same_size()
    {
        AssertFormula.CheckParsingErrorContains("{1,2;3,4;5;6,7}", "Rows of an array don't have same size.");
    }
}