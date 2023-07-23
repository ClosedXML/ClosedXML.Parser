﻿namespace ClosedXML.Parser.Tests.Rules;

/// <summary>
/// Test rule <c>cell_reference</c>.
/// </summary>
[TestClass]
public class CellReferenceRuleTests
{
    [TestMethod]
    [DynamicData(nameof(TokenCombinations))]
    public void Possible_token_combinations_are_converted_to_node(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    public static IEnumerable<object[]> TokenCombinations
    {
        get
        {
            // cell_reference: A1_REFERENCE
            yield return new object[]
            {
                "Z5",
                new LocalReferenceNode(new CellArea(26, 5))
            };

            // "MS-XLSX 2.2.2.1: The formula MUST NOT use the bang-reference or bang-name.
            // cell_reference: BANG_REFERENCE
            // yield return new object[]
            // {
            //     "!A1",
            //     ...
            // };

            // cell_reference: SHEET_RANGE_PREFIX A1_REFERENCE
            yield return new object[]
            {
                "[2]First:Second!B3:D5",
                new ExternalReferenceNode(
                    2,
                    new CellArea(
                        "First",
                        "Second",
                        new CellReference(2, 3),
                        new CellReference(4, 5)))
            };

            // cell_reference: SINGLE_SHEET_PREFIX A1_REFERENCE
            yield return new object[]
            {
                "[2]First!B3:D5",
                new ExternalReferenceNode(
                    2,
                    new CellArea(
                        "First",
                        new CellReference(2, 3),
                        new CellReference(4, 5)))
            };

            // cell_reference: SINGLE_SHEET_PREFIX REF_CONSTANT
            yield return new object[]
            {
                "Sheet5!#REF!",
                new ValueNode("Error", "#REF!")
            };
        }
    }
}
