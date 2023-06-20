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

    private static void AssertText(string formula, string expectedText)
    {
        var node = ParseText(formula, new F());
        Assert.AreEqual("Text", node.Type);
        Assert.AreEqual(expectedText, node.Text);
    }

    private static AstNode ParseText(string formula, IAstFactory<ScalarValue, AstNode> factory)
    {
        var lexer = new FormulaLexer(new CodePointCharStream(formula), TextWriter.Null, TextWriter.Null);
        lexer.RemoveErrorListeners();
        var parser = new FormulaParser<ScalarValue, AstNode>(formula, lexer, factory);
        return parser.Formula();
    }
}