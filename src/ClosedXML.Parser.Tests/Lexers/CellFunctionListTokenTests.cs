namespace ClosedXML.Parser.Tests.Lexers;

public class CellFunctionListTokenTests
{
    [Fact]
    public void Ignores_trailing_whitespaces()
    {
        var expected = new Reference(1, 1);
        Assert.Equal(expected, TokenParser.ExtractCellFunction("A1(  "));
    }

    [Theory]
    [MemberData(nameof(TestData))]
    public void Accepts_absolute_and_relative_cell_addresses(string token, Reference expectedCell)
    {
        Assert.Equal(expectedCell, TokenParser.ExtractCellFunction(token));
    }

    public static IEnumerable<object[]> TestData
    {
        get
        {
            yield return new object[] { "A1(", new Reference(1, 1) };
            yield return new object[] { "$A$1(", new Reference(true, 1, true, 1) };
            yield return new object[] { "$B3(", new Reference(true, 2, false, 3) };
            yield return new object[] { "B$3(", new Reference(false, 2, true, 3) };
        }
    }
}