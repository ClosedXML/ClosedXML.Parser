namespace ClosedXML.Parser.Tests.Rules;

[TestClass]
public class RefIntersectionExpressionRuleTests
{
    [TestMethod]
    [DynamicData(nameof(TestCases))]
    public void Has_one_or_more_elements_separated_by_space(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [TestMethod]
    public void Intersection_operator_has_lower_priority_than_range()
    {
        var expectedNode =
            new BinaryNode(
                BinaryOperation.Intersection,
                new BinaryNode(
                    BinaryOperation.Range,
                    new ValueNode("Error", "#REF!"),
                    new LocalReferenceNode(new CellArea(1, 2))),
                new BinaryNode(
                    BinaryOperation.Range,
                    new LocalReferenceNode(new CellArea(1, 3)),
                    new NameNode("two")));
        AssertFormula.SingleNodeParsed("#REF!:A2 A3:two", expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // ref_intersection_expression : ref_range_expression
            yield return new object[]
            {
                "A1",
                new LocalReferenceNode(new CellArea(1, 1))
            };

            // ref_intersection_expression : ref_range_expression SPACE ref_range_expression
            yield return new object[]
            {
                "A1 A2",
                new BinaryNode(BinaryOperation.Intersection,
                    new LocalReferenceNode(new CellArea(1, 1)),
                    new LocalReferenceNode(new CellArea(1, 2)))
            };

            // ref_intersection_expression : ref_range_expression SPACE ref_range_expression
            yield return new object[]
            {
                " A1   A2   A3  ",
                new BinaryNode(BinaryOperation.Intersection,
                    new BinaryNode(BinaryOperation.Intersection,
                        new LocalReferenceNode(new CellArea(1, 1)),
                        new LocalReferenceNode(new CellArea(1, 2))),
                    new LocalReferenceNode(new CellArea(1, 3)))
            };
        }
    }
}