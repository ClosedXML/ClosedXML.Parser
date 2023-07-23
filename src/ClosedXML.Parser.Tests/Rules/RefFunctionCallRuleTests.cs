namespace ClosedXML.Parser.Tests.Rules;

public class RefFunctionCallRuleTests
{
    [Fact]
    public void Ref_functions_are_recognized()
    {
        var expected = new FunctionNode("IF") { Children = new AstNode[] { new ValueNode("Logical", true) } };
        AssertFormula.SingleNodeParsed("IF(TRUE)", expected);
    }

    [Fact]
    public void Ref_functions_can_have_whitespaces_around_braces()
    {
        var expected = new FunctionNode("CHOOSE") { Children = new AstNode[] { new ValueNode("Logical", true), new ValueNode("Number", 5.0) } };
        AssertFormula.SingleNodeParsed("CHOOSE(  TRUE, 5  )", expected);
    }
}