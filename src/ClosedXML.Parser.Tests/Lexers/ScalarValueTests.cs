using Antlr4.Runtime;

namespace ClosedXML.Parser.Tests.Lexers;

[TestClass]
public class ScalarValueTests
{
    [TestMethod]
    public void Can_parse_empty_string()
    {
        AssertText("\"\"", string.Empty);
    }

    [TestMethod]
    public void Can_parse_non_escaped_string()
    {
        AssertText("\"someone's\ntext\"", "someone's\ntext");
    }

    [TestMethod]
    [DataRow("\"\"\"\"", "\"")]
    [DataRow("\"\"\"\"\"\"", "\"\"")]
    [DataRow("\"Eastern \"\"Bonn's\"\" Tavern\"", "Eastern \"Bonn's\" Tavern")]
    public void Can_parse_escaped_string(string unescaped, string escaped)
    {
        AssertText(unescaped, escaped);
    }

    [TestMethod]
    [DataRow("TRUE", true)]
    [DataRow("FALSE", false)]
    public void Can_parse_logical(string formula, bool value)
    {
        AssertValue(formula, "Logical", value);
    }

    [TestMethod]
    [DataRow("#REF!", "#REF!")]
    [DataRow("#N/A", "#N/A")]
    public void Can_parse_error(string formula, string value)
    {
        AssertValue(formula, "Error", value);
    }

    [TestMethod]
    [DataRow("1", 1)]
    [DataRow("1.5", 1.5)]
    [DataRow(".5", .5)]
    [DataRow(".5E2", 50)]
    // [DataRow(".5e2", 50)] TODO: Lower e
    [DataRow(".5E+2", 50)]
    [DataRow("50E-2", 0.5)]
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
        Assert.AreEqual(expectedType, node.Type);
        Assert.AreEqual(expected, node.Value);
    }

    private static AstNode ParseText(string formula, IAstFactory<ScalarValue, AstNode> factory)
    {
        var lexer = new FormulaLexer(new CodePointCharStream(formula), TextWriter.Null, TextWriter.Null);
        lexer.RemoveErrorListeners();
        var parser = new FormulaParser<ScalarValue, AstNode>(formula, lexer, factory);
        return parser.Formula();
    }
}