namespace ClosedXML.Parser.Tests.Rules;

/// <summary>
/// Test rule <c>cell_reference</c>.
/// </summary>
public class CellReferenceRuleTests
{
    [Theory]
    [MemberData(nameof(TokenCombinationsA1))]
    public void A1_possible_token_combinations_are_converted_to_node(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    public static IEnumerable<object[]> TokenCombinationsA1
    {
        get
        {
            // cell_reference: A1_REFERENCE
            yield return new object[]
            {
                "Z5",
                new ReferenceNode(new ReferenceArea(26, 5))
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
                new ExternalReference3DNode(
                    2,
                    "First",
                    "Second",
                    new ReferenceArea(
                        new Reference(2, 3),
                        new Reference(4, 5)))
            };

            // cell_reference: SINGLE_SHEET_PREFIX A1_REFERENCE
            yield return new object[]
            {
                "[2]First!B3:D5",
                new ExternalReferenceNode(
                    2,
                    new CellArea(
                        "First",
                        new Reference(2, 3),
                        new Reference(4, 5)))
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
