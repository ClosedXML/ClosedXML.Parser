using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests.Rules;

/// <summary>
/// Test rule <c>cell_reference</c>.
/// </summary>
public class CellReferenceRuleTests
{
    private const int MaxCol = 16384;
    private const int MaxRow = 1048576;

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
            // cell_reference: A1_CELL
            yield return new object[]
            {
                "Z5",
                new ReferenceNode(new ReferenceArea(5, 26, A1))
            };

            // cell_reference: A1_CELL COLON A1_CELL
            yield return new object[]
            {
                "A1:A1",
                new ReferenceNode(
                    new ReferenceArea(new RowCol(Relative, 1, Relative, 1, A1)))
            };

            // cell_reference: A1_CELL COLON A1_CELL
            yield return new object[]
            {
                "Z1:AB25",
                new ReferenceNode(
                    new ReferenceArea(
                        new RowCol(1, 26, A1),
                        new RowCol(25, 28, A1)))
            };

            // cell_reference: A1_CELL COLON A1_CELL
            yield return new object[]
            {
                "$XFC$1048575:$XFD$1048576",
                new ReferenceNode(
                    new ReferenceArea(
                        new RowCol(true, MaxRow - 1, true, MaxCol - 1, A1),
                        new RowCol(true, MaxRow, true, MaxCol, A1)))
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
                        new RowCol(3, 2, A1),
                        new RowCol(5, 4, A1)))
                };

            // cell_reference: SINGLE_SHEET_PREFIX A1_CELL COLON A1_CELL
            yield return new object[]
                        {
                "[2]First!B3:D5",
                new ExternalSheetReferenceNode(
                    2,
                    "First",
                    new ReferenceArea(
                        new RowCol(3, 2, A1),
                        new RowCol(5, 4, A1)))
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
