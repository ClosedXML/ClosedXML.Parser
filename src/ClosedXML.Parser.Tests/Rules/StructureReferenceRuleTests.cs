namespace ClosedXML.Parser.Tests.Rules;

[TestClass]
public class StructureReferenceRuleTests
{
    [TestMethod]
    [DynamicData(nameof(TestCases))]
    public void Structure_reference_is_parsed_to_a_node(string formula, AstNode expectedNode)
    {
        AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // structure_reference : INTRA_TABLE_REFERENCE
            yield return new object[]
            {
                "[Column]",
                new StructureReferenceNode(null, StructuredReferenceArea.None, "Column", "Column")
            };

            yield return new object[]
            {
                "[#Totals]",
                new StructureReferenceNode(null, StructuredReferenceArea.Totals, null, null)
            };

            yield return new object[]
            {
                "[]",
                new StructureReferenceNode(null, StructuredReferenceArea.None, null, null)
            };

            yield return new object[]
            {
                "[[#Data],[First Column]:[Last Column]]",
                new StructureReferenceNode(null, StructuredReferenceArea.Data, "First Column", "Last Column")
            };

            // structure_reference : NAME INTRA_TABLE_REFERENCE
            yield return new object[]
            {
                "SomeTable[Column]",
                new StructureReferenceNode("SomeTable", StructuredReferenceArea.None, "Column", "Column")
            };

            // structure_reference: BOOK_PREFIX NAME INTRA_TABLE_REFERENCE
            yield return new object[]
            {
                "[4]!SomeTable[Column]",
                new ExternalStructureReferenceNode(4, "SomeTable", StructuredReferenceArea.None, "Column", "Column")
            };
        }
    }
}