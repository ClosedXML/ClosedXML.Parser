namespace ClosedXML.Parser.Tests.Rules;

public class NameReferenceRuleTests
{
    [Fact]
    public void Name_is_recognized()
    {
        var expectedNode = new NameNode("SomeName");
        AssertFormula.SingleNodeParsed("SomeName", expectedNode);
    }

    [Fact]
    public void Sheet_name_is_recognized()
    {
        var expectedNode = new SheetNameNode("Sheet", "SomeName");
        AssertFormula.SingleNodeParsed("Sheet!SomeName", expectedNode);
    }

    [Fact]
    public void External_name_is_recognized()
    {
        var expectedNode = new ExternalNameNode(2, "SomeName");
        AssertFormula.SingleNodeParsed("[2]!SomeName", expectedNode);
    }

    [Fact]
    public void External_sheet_name_is_recognized()
    {
        var expectedNode = new ExternalSheetNameNode(14, "Sheet", "SomeName");
        AssertFormula.SingleNodeParsed("[14]Sheet!SomeName", expectedNode);
    }
}