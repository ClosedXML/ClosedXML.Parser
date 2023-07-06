namespace ClosedXML.Parser.Tests.Rules;

/// <summary>
/// Test rule <c>cell_reference</c>.
/// </summary>
[TestClass]
public class CellReferenceRuleTests
{
    // cell_reference: A1_REFERENCE
    [TestMethod]
    public void A1_REFERENCE_extracts_local_reference()
    {
        var expected = new LocalReferenceNode(new CellArea(26, 5));
        AssertFormula.SingleNodeParsed("Z5", expected);
    }

    // cell_reference: BANG_REFERENCE
    [TestMethod]
    [Ignore("MS-XLSX 2.2.2.1: The formula MUST NOT use the bang-reference or bang-name.")]
    public void BANG_REFERENCE_is_external_cell_reference()
    {
        Assert.Fail("TODO");
    }

    // cell_reference: SHEET_RANGE_PREFIX A1_REFERENCE
    [TestMethod]
    public void SHEET_RANGE_PREFIX__A1_REFERENCE_is_external_cell_reference()
    {
        var node = new ExternalReferenceNode(
            2,
            new CellArea(
                "First",
                "Second",
                new CellReference(2, 3),
                new CellReference(4, 5)));
        AssertFormula.SingleNodeParsed("[2]First:Second!B3:D5", node);
    }

    [TestMethod]
    public void SINGLE_SHEET_PREFIX__A1_REFERENCE_is_external_cell_reference()
    {
        Assert.Fail("TODO");
    }

    [TestMethod]
    public void SINGLE_SHEET_PREFIX__REF_CONSTANT_is_external_cell_reference()
    {
        Assert.Fail("TODO");
    }
}