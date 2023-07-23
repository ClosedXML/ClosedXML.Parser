namespace ClosedXML.Parser.Tests.Lexers;

public class CellFunctionListTokenTests
{
    [Fact]
    public void Ignores_trailing_whitespaces()
    {
        var expected = new CellReference(1, 1);
        Assert.Equal(expected, TokenParser.ExtractCellFunction("A1(  "));
    }

    [Theory]
    [MemberData(nameof(TestData))]
    public void Accepts_absolute_and_relative_cell_addresses(string token, CellReference expectedCell)
    {
        Assert.Equal(expectedCell, TokenParser.ExtractCellFunction(token));
    }

    public static IEnumerable<object[]> TestData
    {
        get
        {
            yield return new object[] { "A1(", new CellReference(1, 1) };
            yield return new object[] { "$A$1(", new CellReference(true, 1, true, 1) };
            yield return new object[] { "$B3(", new CellReference(true, 2, false, 3) };
            yield return new object[] { "B$3(", new CellReference(false, 2, true, 3) };
        }
    }
}