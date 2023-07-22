namespace ClosedXML.Parser.Tests.Rules;

[TestClass]
public class RefAtomExpressionRuleTests
{
    [TestMethod]
    public void Ref_error()
    {
        VerifyNode("#REF!", new ValueNode("Error", "#REF!"));
    }

    [TestMethod]
    public void Cell_reference()
    {
        VerifyNode("A1", new LocalReferenceNode(new CellArea(1, 1)));
    }

    [TestMethod]
    public void Name_reference()
    {
        VerifyNode("some_name", new NameNode("some_name"));
    }

    [TestMethod]
    public void Structure_reference()
    {
        VerifyNode("Table[Column]", new StructureReferenceNode("Table", StructuredReferenceArea.None, "Column", "Column"));
    }

    [TestMethod]
    public void Nested_ref_expression()
    {
        VerifyNode("((#REF!))", new ValueNode("Error", "#REF!"));
    }

    [TestMethod]
    public void Nested_cant_be_non_ref()
    {
        AssertFormula.CheckParsingErrorContains("(1),#REF!", "not completely parsed");
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