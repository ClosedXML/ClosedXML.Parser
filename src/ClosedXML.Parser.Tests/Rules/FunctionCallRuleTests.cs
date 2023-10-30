namespace ClosedXML.Parser.Tests.Rules;

public class FunctionCallRuleTests
{
    [Fact]
    public void Predefined_functions_are_recognized()
    {
        var expectedNode = new FunctionNode("SIN") { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("SIN(5)", expectedNode);
    }

    [Fact]
    public void Function_can_have_whitespaces_around_braces()
    {
        var expectedNode = new FunctionNode("SIN") { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("SIN(  5  )", expectedNode);
    }

    [Fact]
    public void Function_can_be_from_another_sheet()
    {
        var expectedNode = new FunctionNode("Sheet", "Func") { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("Sheet!Func(5)", expectedNode);
    }

    [Fact]
    public void Function_can_be_from_another_workbook()
    {
        var expectedNode = new ExternalFunctionNode(2, null, "Func") { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("[2]!Func(5)", expectedNode);
    }

    [Fact]
    public void Function_can_be_cell_function()
    {
        var expectedNode = new CellFunctionNode(new RowCol(true, 3, false, 2, A1)) { Children = new AstNode[] { new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("B$3(5)", expectedNode);
    }
}