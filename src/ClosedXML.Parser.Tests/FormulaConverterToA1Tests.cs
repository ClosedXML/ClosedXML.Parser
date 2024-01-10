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

    [Theory]
    [InlineData("R[-4]C", 4, 1, "#REF!")]
    [InlineData("R[1048575]C", 2, 1, "#REF!")]
    [InlineData("RC[-4]", 1, 4, "#REF!")]
    [InlineData("RC[5]", 1, 16380, "#REF!")]
    public void Out_of_bounds_references(string r1c1, int row, int col, string a1)
    {
        Assert.Equal(a1, FormulaConverter.ToA1(r1c1, row, col));
    }

    [Theory]
    [InlineData("C2:C4", 1, 1, "$B:$D")]
    [InlineData("C[-2]:C4", 1, 4, "B:$D")]
    [InlineData("C2:C[3]", 1, 4, "$B:G")]
    [InlineData("C[2]:C[3]", 1, 4, "F:G")]
    public void Columns_reference(string r1c1, int row, int col, string a1)
    {
        Assert.Equal(a1, FormulaConverter.ToA1(r1c1, row, col));
    }

    [Theory]
    [InlineData("R2:R4", 1, 1, "$2:$4")]
    [InlineData("R[-2]:R4", 4, 1, "2:$4")]
    [InlineData("R2:R[3]", 4, 1, "$2:7")]
    [InlineData("R[2]:R[3]", 4, 1, "6:7")]
    public void Rows_reference(string r1c1, int row, int col, string a1)
    {
        Assert.Equal(a1, FormulaConverter.ToA1(r1c1, row, col));
    }

}