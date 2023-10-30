namespace ClosedXML.Parser.Tests;

/// <summary>
/// Most of tests are taken care of by <see cref="FormulaConverterToR1C1Tests"/>.
/// </summary>
public class FormulaConverterToA1Tests
{
    [Theory]
    [InlineData("RC", 4, 1, "A4")] // References
    [InlineData("R7C", 4, 2, "B$7")]
    [InlineData("R[2]C", 4, 2, "B6")]
    [InlineData("RC5", 4, 2, "$E4")]
    [InlineData("RC[3]", 4, 2, "E4")]
    [InlineData("R[2]C[1]:R[3]C[4]", 4, 3, "D6:G7")]
    [InlineData("C[-2]:C[1]", 4, 3, "A:D")]
    [InlineData("R[-2]:R[6]", 4, 3, "2:10")]
    [InlineData("Sheet4!R[2]C", 4, 2, "Sheet4!B6")] // Sheet reference
    [InlineData("R7C3(TRUE)", 4, 2, "$E$11(TRUE)", Skip = "Parser bug")] // Cell function
    public void ExternalSheetReference(string r1c1, int row, int col, string a1)
    {
        Assert.Equal(a1, FormulaConverter.ToA1(r1c1, row, col));
    }
}