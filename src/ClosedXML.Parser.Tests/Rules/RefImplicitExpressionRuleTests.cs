namespace ClosedXML.Parser.Tests.Rules;

public class RefImplicitExpressionRuleTests
{
    [Fact]
    public void Implicit_intersection_operator_has_lower_priority_than_range()
    {
        var expectedNode =
            new UnaryNode(
                UnaryOperation.ImplicitIntersection,
                new BinaryNode(
                    BinaryOperation.Range,
                    new ReferenceNode(new ReferenceArea(new RowCol(1, 1, A1), new RowCol(2, 1, A1))),
                    new ReferenceNode(new ReferenceArea(3, 1, A1))));
        AssertFormula.SingleNodeParsed("@A1:A2:A3", expectedNode);
    }
}