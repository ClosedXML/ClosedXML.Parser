using ClosedXML.Parser.Pratt;

namespace ClosedXML.Parser.Tests.Lexers;

public class ParseletIdentTests
{
    [Theory]

    // Local area
    [InlineData("A1:B1", typeof(ReferenceNode))]
    [InlineData("$A$1:$B$1", typeof(ReferenceNode))]

    // Local cell
    [InlineData("A1", typeof(ReferenceNode))]
    [InlineData("A$1", typeof(ReferenceNode))]
    [InlineData("$A1", typeof(ReferenceNode))]
    [InlineData("$A$1", typeof(ReferenceNode))]
    [InlineData("XFD1048576", typeof(ReferenceNode))]
    [InlineData("XFD$1048576", typeof(ReferenceNode))]
    [InlineData("$XFD1048576", typeof(ReferenceNode))]
    [InlineData("$XFD$1048576", typeof(ReferenceNode))]

    // Local colspan
    [InlineData("A:B", typeof(ReferenceNode))]
    [InlineData("$GE:$XFD", typeof(ReferenceNode))]

    // Local rowspan starting with absolute
    [InlineData("$1:8", typeof(ReferenceNode))]
    [InlineData("$72:$85", typeof(ReferenceNode))]

    // sheet!A1:A2
    [InlineData("Sheet!A1:B2", typeof(SheetReferenceNode))]
    [InlineData("Sheet!$Z$84:$BG$99", typeof(SheetReferenceNode))]

    // sheet!A1
    [InlineData("Sheet!A1", typeof(SheetReferenceNode))]
    [InlineData("Sheet!$Z$84", typeof(SheetReferenceNode))]

    // sheet!$1:2
    [InlineData("Sheet!$4:81", typeof(SheetReferenceNode))]
    [InlineData("Sheet!$1:$5", typeof(SheetReferenceNode))]

    // sheet!name
    [InlineData("Sheet!name", typeof(SheetNameNode))]
    [InlineData("Sheet!_name", typeof(SheetNameNode))]

    // sheet!1:2
    [InlineData("Sheet!1:2", typeof(SheetReferenceNode))]
    [InlineData("Sheet!1:$2", typeof(SheetReferenceNode))]

    // name
    [InlineData("_name", typeof(NameNode))]
    [InlineData("name", typeof(NameNode))]
    public void Can_parse_references_starting_at_ident(string formula, Type expectedNodeType)
    {
        var parser = ParserFactory.Create(new F());
        var root = parser.ParseFormula(formula, new Ctx());

        Assert.Equal(expectedNodeType, root.GetType());
        Assert.Equal(formula, root.GetDisplayString(A1));
    }

    [Theory]
    [InlineData("TRUE", true)]
    [InlineData("true", true)]
    [InlineData("FALSE", false)]
    [InlineData("false", false)]
    public void Can_parse_logical(string formula, bool expectedValue)
    {
        var parser = ParserFactory.Create(new F());
        var root = parser.ParseFormula(formula, new Ctx());

        Assert.Equal(new ValueNode(expectedValue), root);
    }

    [Theory]
    [InlineData("sheet!$")]
    [InlineData("sheet!")]
    [InlineData("$")]
    [InlineData("A01")]
    [InlineData("A0")]
    [InlineData("A1048577")]
    [InlineData("XFE1")]
    public void Invalid_references_starting_with_ident_throw_parsing_exception(string formula)
    {
        var parser = ParserFactory.Create(new F());
        Assert.Throws<ParsingException>(() => parser.ParseFormula(formula, new Ctx()));
    }
}
