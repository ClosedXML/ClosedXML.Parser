﻿namespace ClosedXML.Parser.Tests.Rules;

public class RefExpressionRuleTests
{
    [Theory]
    [MemberData(nameof(TestCases))]
    public void Ref_expression_can_have_multiple_ref_intersection_expressions(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Fact]
    public void Union_operator_has_lower_priority_than_intersection()
    {
        var expectedNode =
            new BinaryNode(
                BinaryOperation.Union,
                new BinaryNode(
                    BinaryOperation.Intersection,
                    new ReferenceNode(new ReferenceArea(1, 1, A1)),
                    new ReferenceNode(new ReferenceArea(2, 1, A1))),
                new BinaryNode(
                    BinaryOperation.Intersection,
                    new ReferenceNode(new ReferenceArea(3, 1, A1)),
                    new ReferenceNode(new ReferenceArea(4, 1, A1))));
        AssertFormula.SingleNodeParsed("A1 A2,A3 A4", expectedNode);
    }

    [Fact]
    public void Whitespaces_at_the_end_of_formula_are_ignored()
    {
        var expectedNode = new BinaryNode(
            BinaryOperation.Union,
            new NameNode("some_name"),
            new ReferenceNode(new ReferenceArea(2, 1, A1)));
        AssertFormula.SingleNodeParsed(" some_name , A2 ", expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // ref_expression : ref_intersection_expression
            yield return new object[]
            {
                "A1",
                new ReferenceNode(new ReferenceArea(1, 1, A1))
            };

            // ref_expression : ref_intersection_expression COMMA ref_intersection_expression
            yield return new object[]
            {
                "A1,A2",
                new BinaryNode(
                    BinaryOperation.Union,
                    new ReferenceNode(new ReferenceArea(1, 1, A1)),
                    new ReferenceNode(new ReferenceArea(2, 1, A1)))
            };

            // ref_expression : ref_intersection_expression COMMA ref_intersection_expression
            yield return new object[]
            {
                "A1,#REF!,A2",
                new BinaryNode(
                    BinaryOperation.Union,
                    new BinaryNode(BinaryOperation.Union,
                        new ReferenceNode(new ReferenceArea(1, 1, A1)),
                        new ValueNode("Error", "#REF!")),
                    new ReferenceNode(new ReferenceArea(2, 1, A1)))
            };
        }
    }
}