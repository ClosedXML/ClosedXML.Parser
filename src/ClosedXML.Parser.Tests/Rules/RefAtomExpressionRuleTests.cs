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
        var adjustedFormula = $"{formula},#REF!";
        var expected = new BinaryNode(BinaryOperation.Union, node, new ValueNode("Error", "#REF!"));
        AssertFormula.SingleNodeParsed(adjustedFormula, expected);
    }
}