﻿using Microsoft.VisualStudio.TestPlatform.ObjectModel;

namespace ClosedXML.Parser.Tests.Rules;

[TestClass]
public class RefExpressionRuleTests
{
    [TestMethod]
    [DynamicData(nameof(TestCases))]
    public void Ref_expression_can_have_multiple_ref_intersection_expressions(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [TestMethod]
    public void Union_operator_has_lower_priority_than_intersection()
    {
        var expectedNode =
            new BinaryNode(
                BinaryOperation.Union,
                new BinaryNode(
                    BinaryOperation.Intersection,
                    new LocalReferenceNode(new CellArea(1, 1)),
                    new LocalReferenceNode(new CellArea(1, 2))),
                new BinaryNode(
                    BinaryOperation.Intersection,
                    new LocalReferenceNode(new CellArea(1, 3)),
                    new LocalReferenceNode(new CellArea(1, 4))));
        AssertFormula.SingleNodeParsed("A1 A2,A3 A4", expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // ref_expression : ref_intersection_expression
            yield return new object[]
            {
                "A1",
                new LocalReferenceNode(new CellArea(1,1))
            };

            // ref_expression : ref_intersection_expression COMMA ref_intersection_expression
            yield return new object[]
            {
                "A1,A2",
                new BinaryNode(
                    BinaryOperation.Union,
                    new LocalReferenceNode(new CellArea(1,1)),
                    new LocalReferenceNode(new CellArea(1,2)))
            };

            // ref_expression : ref_intersection_expression COMMA ref_intersection_expression
            yield return new object[]
            {
                "A1,#REF!,A2",
                new BinaryNode(
                    BinaryOperation.Union,
                    new BinaryNode(BinaryOperation.Union,
                        new LocalReferenceNode(new CellArea(1,1)),
                        new ValueNode("Error", "#REF!")),
                    new LocalReferenceNode(new CellArea(1,2)))
            };
        }
    }

}