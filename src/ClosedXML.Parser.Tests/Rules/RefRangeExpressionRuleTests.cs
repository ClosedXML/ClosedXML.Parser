namespace ClosedXML.Parser.Tests.Rules;

public class RefRangeExpressionRuleTests
{
    [Theory]
    [MemberData(nameof(TestData))]
    public void Has_one_or_more_elements_separated_by_space(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    public static IEnumerable<object[]> TestData
    {
        get
        {
            // ref_range_expression : ref_atom_expression
            yield return new object[]
            {
                "A1",
                new ReferenceNode(new CellArea(1, 1))
            };

            // ref_range_expression : ref_atom_expression COLON ref_atom_expression
            yield return new object[]
            {
                "first:second",
                new BinaryNode(
                    BinaryOperation.Range,
                    new NameNode("first"),
                    new NameNode("second"))
            };

            // ref_range_expression : ref_atom_expression COLON ref_atom_expression COLON ref_atom_expression
            yield return new object[]
            {
                // Parser eats A1:B2 as a single token
                "#REF!:B1:last",
                new BinaryNode(
                    BinaryOperation.Range,
                    new BinaryNode(
                        BinaryOperation.Range,
                        new ValueNode("Error", "#REF!"),
                        new ReferenceNode(new CellArea(2, 1))),
                    new NameNode("last"))
            };
        }
    }
}