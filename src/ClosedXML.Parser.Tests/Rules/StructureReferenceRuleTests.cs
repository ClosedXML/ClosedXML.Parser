namespace ClosedXML.Parser.Tests.Rules;

public class StructureReferenceRuleTests
{
    [Theory]
    [MemberData(nameof(TestCases))]
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

            // A keyword list is a whole inner reference on its own, with no column range after it.
            yield return new object[]
            {
                "[[#All]]",
                new StructureReferenceNode(null, StructuredReferenceArea.All, null, null)
            };

            yield return new object[]
            {
                "[[#Headers],[#Data]]",
                new StructureReferenceNode(null, StructuredReferenceArea.Headers | StructuredReferenceArea.Data, null, null)
            };

            // structure_reference : NAME INTRA_TABLE_REFERENCE
            yield return new object[]
            {
                "SomeTable[Column]",
                new StructureReferenceNode("SomeTable", StructuredReferenceArea.None, "Column", "Column")
            };

            yield return new object[]
            {
                "SomeTable[[#Data],[#Totals]]",
                new StructureReferenceNode("SomeTable", StructuredReferenceArea.Data | StructuredReferenceArea.Totals, null, null)
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