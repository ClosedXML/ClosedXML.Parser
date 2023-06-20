using Antlr4.Runtime;
using ClosedXML.Lexer;

namespace ClosedXML.Parser.Tests;

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

    private static void AssertText<T>(string formula, T expected)
    {
        AssertValue(formula, "Text", expected);
    }

    private static void AssertValue<T>(string formula, string expectedType, T expected)
    {
        var node = ParseText(formula, new F());
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