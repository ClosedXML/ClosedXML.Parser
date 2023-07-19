namespace ClosedXML.Parser.Tests.Rules;

[TestClass]
public class FunctionCallRuleTests
{
    [TestMethod]
    public void Predefined_functions_are_recognized()
    {
        var expected = new FunctionNode("SIN") { Children = new AstNode[] { new ValueNode("Number", 5.0) }};
        AssertFormula.SingleNodeParsed("SIN(5)", expected);
    }

    [TestMethod]
    public void Function_can_have_whitespaces_around_braces()
    {
        var expected = new FunctionNode("SIN") { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("SIN(  5  )", expected);
    }

    [TestMethod]
    public void Function_can_be_from_another_sheet()
    {
        var expected = new FunctionNode("Sheet", "Func") { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("Sheet!Func(5)", expected);
    }

    [TestMethod]
    public void Function_can_be_from_another_workbook()
    {
        var expected = new ExternalFunctionNode(2, null, "Func") { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("[2]!Func(5)", expected);
    }
}