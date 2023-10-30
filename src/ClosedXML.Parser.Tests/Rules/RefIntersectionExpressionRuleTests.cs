namespace ClosedXML.Parser.Tests.Rules;

public class RefIntersectionExpressionRuleTests
{
    [Theory]
    [MemberData(nameof(TestCases))]
    public void Has_one_or_more_elements_separated_by_space(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Fact]
    public void Intersection_operator_has_lower_priority_than_implicit_intersection()
    {
        var expectedNode =
            new UnaryNode(
                UnaryOperation.ImplicitIntersection,
                new BinaryNode(
                    BinaryOperation.Intersection,
                    new ReferenceNode(new ReferenceArea(new RowCol(1,1, A1), new RowCol(10,1, A1))),
                    new ReferenceNode(new ReferenceArea(5, 1, A1))));
        AssertFormula.SingleNodeParsed("@A1:A10 A5", expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // ref_intersection_expression : ref_range_expression
            yield return new object[]
            {
                "A1",
                new ReferenceNode(new ReferenceArea(1, 1, A1))
            };

            // ref_intersection_expression : ref_range_expression SPACE ref_range_expression
            yield return new object[]
            {
                "A1 A2",
                new BinaryNode(BinaryOperation.Intersection,
                    new ReferenceNode(new ReferenceArea(1, 1, A1)),
                    new ReferenceNode(new ReferenceArea(2, 1, A1)))
            };

            // ref_intersection_expression : ref_range_expression SPACE ref_range_expression
            yield return new object[]
            {
                " A1   A2   A3  ",
                new BinaryNode(BinaryOperation.Intersection,
                    new BinaryNode(BinaryOperation.Intersection,
                        new ReferenceNode(new ReferenceArea(1, 1, A1)),
                        new ReferenceNode(new ReferenceArea(2, 1, A1))),
                    new ReferenceNode(new ReferenceArea(3, 1, A1)))
            };
        }
    }
}