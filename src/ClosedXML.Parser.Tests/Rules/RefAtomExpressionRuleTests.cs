namespace ClosedXML.Parser.Tests.Rules;

public class RefAtomExpressionRuleTests
{
    [Fact]
    public void Ref_error()
    {
        VerifyNode("#REF!", new ValueNode("Error", "#REF!"));
    }

    [Fact]
    public void Nested_ref_expression()
    {
        VerifyNode("((#REF!))", new ValueNode("Error", "#REF!"));
    }

    [Fact]
    public void Cell_reference()
    {
        VerifyNode("A1", new ReferenceNode(new ReferenceArea(1, 1, A1)));
    }

    [Fact]
    public void Ref_function_call()
    {
        VerifyNode("IF(TRUE,B5)", new FunctionNode("IF")
        {
            Children = new AstNode[]
            {
                new ValueNode(true),
                new ReferenceNode(new ReferenceArea(5, 2, A1))
            }
        });
    }

    [Fact]
    public void Name_reference()
    {
        VerifyNode("some_name", new NameNode("some_name"));
    }

    [Fact]
    public void Name_reference_starts_3D_reference()
    {
        VerifyNode("Sheet1:Sheet3!A1", new Reference3DNode("Sheet1", "Sheet3", new ReferenceArea(1, 1, A1)));
    }

    [Fact]
    public void Sheet_reference_can_have_whitespace_after_exclamation_mark()
    {
        VerifyNode("Sheet2!  A1", new SheetReferenceNode("Sheet2", new ReferenceArea(1, 1, A1)));
    }

    [Fact]
    public void Structure_reference()
    {
        VerifyNode("Table[Column]", new StructureReferenceNode("Table", StructuredReferenceArea.None, "Column", "Column"));
    }

    [Fact]
    public void Nested_cant_be_non_ref()
    {
        AssertFormula.CheckParsingErrorContains("(1),#REF!", "The formula `(1),#REF!` wasn't parsed correctly. The expression `(1)` was parsed, but the rest `,#REF!` wasn't.");
    }

    [Fact]
    public void Non_ref_function_cant_be_ref_atom()
    {
        AssertFormula.CheckParsingErrorContains("FUNC(),#REF!", "The formula `FUNC(),#REF!` wasn't parsed correctly. The expression `FUNC()` was parsed, but the rest `,#REF!` wasn't.");
    }

    private static void VerifyNode(string formula, AstNode node)
    {
        // Force the formula into a ref_atom_expression through ref_range_expression. 
        // ref_range_expression: ref_atom_expression (COLON ref_atom_expression)*
        var adjustedFormula = $"#REF!:{formula}";
        var expected = new BinaryNode(BinaryOperation.Range, new ValueNode("Error", "#REF!"), node);
        AssertFormula.SingleNodeParsed(adjustedFormula, expected);
    }
}