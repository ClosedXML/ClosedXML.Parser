
namespace ClosedXML.Parser.Tests.Rules;

[TestClass]
public class ConstantRuleTests
{
    [TestMethod]
    [DataRow("#DIV/0!")]
    [DataRow("#N/A")]
    [DataRow("#NAME?")]
    [DataRow("#NULL!")]
    [DataRow("#NUM!")]
    [DataRow("#VALUE!")]
    [DataRow("#GETTING_DATA")]
    public void NonRefErrors(string error)
    {
        AssertFormula.SingleNodeParsed(error, new ValueNode("Error", error));
    }

    [TestMethod]
    [DataRow("TRUE", true)]
    [DataRow("FALSE", false)]
    [DataRow("true", true)]
    [DataRow("false", false)]
    public void LogicalConstant(string formula, bool value)
    {
        AssertFormula.SingleNodeParsed(formula, new ValueNode("Logical", value));
    }

    [TestMethod]
    [DataRow("1.5e2", 150.0)]
    [DataRow("25.0e-2", 0.25)]
    [DataRow("1", 1.0)]
    [DataRow("5.4", 5.4)]
    public void NumericalConstant(string formula, double value)
    {
        AssertFormula.SingleNodeParsed(formula, new ValueNode("Number", value));
    }

    [TestMethod]
    [DataRow("\"Hello\"", "Hello")]
    [DataRow("\"Tom \"\"Ben\"\"\"", "Tom \"Ben\"")]
    [DataRow("\"\"", "")]
    public void StringConstant(string formula, string text)
    {
        AssertFormula.SingleNodeParsed(formula, new ValueNode("Text", text));
    }

    [TestMethod]
    public void Single_element_array()
    {
        AssertFormula.SingleNodeParsed("{1}", new ArrayNode(1, 1, new[] { new ScalarValue(1) }));
    }

    [TestMethod]
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

    [TestMethod]
    public void Array_cant_contain_blanks()
    {
        AssertFormula.CheckParsingErrorContains("{1,,}", " Unexpected token COMMA.");
    }

    [TestMethod]
    public void Empty_array_is_unparsable()
    {
        AssertFormula.CheckParsingErrorContains("{}", "Unexpected token CLOSE_CURLY.");
    }

    [TestMethod]
    public void Rows_of_array_must_have_same_size()
    {
        AssertFormula.CheckParsingErrorContains("{1,2;3,4;5;6,7}", "Rows of an array don't have same size.");
    }
}