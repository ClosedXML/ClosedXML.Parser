namespace ClosedXML.Parser.Tests.Rules;

[TestClass]
internal class RefFunctionCallTests
{
    [TestMethod]
    public void Ref_functions_are_recognized()
    {
        var expected = new FunctionNode("IF") { Children = new AstNode[] { new ValueNode("Logical", true) } };
        AssertFormula.SingleNodeParsed("IF(TRUE)", expected);
    }
}