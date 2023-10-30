namespace ClosedXML.Parser.Tests.Lexers;

public class CellFunctionListTokenTests
{
    [Fact]
    public void Ignores_trailing_whitespaces()
    {
        var expected = new RowCol(1, 1, A1);
        Assert.Equal(expected, TokenParser.ExtractCellFunction("A1(  "));
    }

    [Theory]
    [MemberData(nameof(TestData))]
    public void Accepts_absolute_and_relative_cell_addresses(string token, RowCol expectedCell)
    {
        Assert.Equal(expectedCell, TokenParser.ExtractCellFunction(token));
    }

    public static IEnumerable<object[]> TestData
    {
        get
        {
            yield return new object[] { "A1(", new RowCol(1, 1, A1) };
            yield return new object[] { "$A$1(", new RowCol(true, 1, true, 1, A1) };
            yield return new object[] { "$B3(", new RowCol(false, 3, true, 2, A1) };
            yield return new object[] { "B$3(", new RowCol(true, 3, false, 2, A1) };
        }
    }
}