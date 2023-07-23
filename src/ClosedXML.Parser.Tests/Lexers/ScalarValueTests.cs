using Xunit;

namespace ClosedXML.Parser.Tests.Lexers;

public class ScalarValueTests
{
    [Fact]
    public void Can_parse_empty_string()
    {
        AssertText("\"\"", string.Empty);
    }

    [Fact]
    public void Can_parse_non_escaped_string()
    {
        AssertText("\"someone's\ntext\"", "someone's\ntext");
    }

    [Theory]
    [InlineData("\"\"\"\"", "\"")]
    [InlineData("\"\"\"\"\"\"", "\"\"")]
    [InlineData("\"Eastern \"\"Bonn's\"\" Tavern\"", "Eastern \"Bonn's\" Tavern")]
    public void Can_parse_escaped_string(string unescaped, string escaped)
    {
        AssertText(unescaped, escaped);
    }

    [Theory]
    [InlineData("TRUE", true)]
    [InlineData("FALSE", false)]
    public void Can_parse_logical(string formula, bool value)
    {
        AssertValue(formula, "Logical", value);
    }

    [Theory]
    [InlineData("#REF!", "#REF!")]
    [InlineData("#N/A", "#N/A")]
    public void Can_parse_error(string formula, string value)
    {
        AssertValue(formula, "Error", value);
    }

    [Theory]
    [InlineData("1", 1)]
    [InlineData("1.5", 1.5)]
    [InlineData(".5", .5)]
    [InlineData(".5E2", 50)]
    // [InlineData(".5e2", 50)] TODO: Lower e
    [InlineData(".5E+2", 50)]
    [InlineData("50E-2", 0.5)]
    public void Can_parse_number(string formula, double value)
    {
        AssertValue(formula, "Number", value);
    }

    private static void AssertText<T>(string formula, T expected)
    {
        AssertValue(formula, "Text", expected);
    }

    private static void AssertValue<T>(string formula, string expectedType, T expected)
    {
        var node = (ValueNode)ParseText(formula, new F());
        Assert.Equal(expectedType, node.Type);
        Assert.Equal(expected, node.Value);
    }

    private static AstNode ParseText(string formula, IAstFactory<ScalarValue, AstNode> factory)
    {
        var parser = new FormulaParser<ScalarValue, AstNode>(formula, factory);
        return parser.Formula();
    }
}