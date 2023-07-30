namespace ClosedXML.Parser.Tests.Rules;

public class PercentExpressionRuleTests
{
    [Fact]
    public void There_can_be_multiple_percent_operators()
    {
        var expectedNode = new UnaryNode(UnaryOperation.Percent)
        {
            Children = new AstNode[]
            {
                new UnaryNode(UnaryOperation.Percent)
                {
                    Children = new AstNode[]
                    {
                        new ValueNode(1234)
                    }
                }
            }
        };
        AssertFormula.SingleNodeParsed("1234%%", expectedNode);
    }
}