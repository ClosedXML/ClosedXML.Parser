namespace ClosedXML.Parser.Tests.Rules;

public class PrefixAtomExpressionRuleTests
{
    [Theory]
    [InlineData("++1", UnaryOperation.Plus)]
    [InlineData("--1", UnaryOperation.Minus)]
    [InlineData("@@1", UnaryOperation.ImplicitIntersection)]
    public void Multiple_unary_operators(string formula, UnaryOperation op)
    {
        var expectedNode = new UnaryNode(op)
        {
            Children = new AstNode[]
            {
                new UnaryNode(op)
                {
                    Children = new AstNode[]
                    {
                        new ValueNode(1.0)
                    }
                }
            }
        };

        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }
}