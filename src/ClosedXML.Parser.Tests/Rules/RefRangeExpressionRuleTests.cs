namespace ClosedXML.Parser.Tests.Rules;

public class RefRangeExpressionRuleTests
{
    [Theory]
    [MemberData(nameof(TestData))]
    public void Has_one_or_more_elements_separated_by_space(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Fact]
    public void Cell_are_not_mistakenly_recognized_as_3d_reference()
    {
        var formula = "A1:Sheet1!B2";
        var expectedNode = new BinaryNode(BinaryOperation.Range)
        {
            Children = new AstNode[]
            {
                new ReferenceNode(new ReferenceSymbol(1, 1)),
                new SheetReferenceNode("Sheet1", new ReferenceSymbol(2, 2))
            }
        };
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Fact]
    public void Columns_can_be_sheet_names_for_3d_reference()
    {
        // JAN and DEC are columns in A1 notation
        var formula = "JAN:DEC!B2";
        var expectedNode = new Reference3DNode("JAN", "DEC", new ReferenceSymbol(2, 2));
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Fact]
    public void Spill_has_higher_priority_than_range()
    {
        var expectedNode = new BinaryNode(BinaryOperation.Range)
        {
            Children = new AstNode[]
            {
                new ReferenceNode(new ReferenceSymbol(5, 1)),
                new UnaryNode(UnaryOperation.SpillRange)
                {
                    Children = new AstNode[]
                    {
                        new NameNode("Name")
                    }
                }
            }
        };

        AssertFormula.SingleNodeParsed("A5:Name#", expectedNode);
    }

    public static IEnumerable<object[]> TestData
    {
        get
        {
            // ref_range_expression : ref_atom_expression
            yield return new object[]
            {
                "A1",
                new ReferenceNode(new ReferenceSymbol(1, 1))
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
                        new ReferenceNode(new ReferenceSymbol(1, 2))),
                    new NameNode("last"))
            };

            // ref_range_expression : A1_CELL COLON NAME COLON A1_CELL
            yield return new object[]
            {
                "A5:B6C7:D8", // Make sure parser doesn't mistake first part of formula for area A5:B6
                new BinaryNode(
                    BinaryOperation.Range,
                    new BinaryNode(
                        BinaryOperation.Range,
                        new ReferenceNode(new ReferenceSymbol(5, 1)),
                        new NameNode("B6C7")),
                    new ReferenceNode(new ReferenceSymbol(8, 4)))
            };
        }
    }
}