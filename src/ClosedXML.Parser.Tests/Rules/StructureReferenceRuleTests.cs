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
                new StructureReferenceNode(null, StructuredReferenceSpecific.None, "Column", "Column")
            };

            yield return new object[]
            {
                "[#Totals]",
                new StructureReferenceNode(null, StructuredReferenceSpecific.Totals, null, null)
            };

            yield return new object[]
            {
                "[]",
                new StructureReferenceNode(null, StructuredReferenceSpecific.None, null, null)
            };

            yield return new object[]
            {
                "[[#Data],[First Column]:[Last Column]]",
                new StructureReferenceNode(null, StructuredReferenceSpecific.Data, "First Column", "Last Column")
            };

            // structure_reference : NAME INTRA_TABLE_REFERENCE
            yield return new object[]
            {
                "SomeTable[Column]",
                new StructureReferenceNode("SomeTable", StructuredReferenceSpecific.None, "Column", "Column")
            };

            // structure_reference: BOOK_PREFIX NAME INTRA_TABLE_REFERENCE
            yield return new object[]
            {
                "[4]!SomeTable[Column]",
                new ExternalStructureReferenceNode(4, "SomeTable", StructuredReferenceSpecific.None, "Column", "Column")
            };
        }
    }
}