using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests;

public class RowColTests
{
    [Fact]
    public void Default_struct_is_A1()
    {
        RowCol a = default;

        Assert.Equal(Relative, a.ColumnType);
        Assert.Equal(1, a.ColumnValue);
        Assert.Equal(Relative, a.RowType);
        Assert.Equal(1, a.RowValue);
        Assert.Equal(A1, a.Style);
    }

    [Theory]
    [InlineData("RC", 1, 1, "A1")]
    [InlineData("RC[-5]", 1, 4, "XFC1")]
    [InlineData("RC[-4]", 1, 4, "XFD1")]
    [InlineData("RC[-3]", 1, 4, "A1")]
    [InlineData("RC[2]", 1, 16382, "XFD1")]
    [InlineData("RC[3]", 1, 16382, "A1")]
    [InlineData("RC[4]", 1, 16382, "B1")]
    [InlineData("R[0]C", 1, 1, "A1")]
    [InlineData("R[-3]C", 4, 1, "A1")]
    [InlineData("R[-4]C", 4, 1, "A1048576")]
    [InlineData("R[-5]C", 4, 1, "A1048575")]
    [InlineData("R[1]C", 1048575, 1, "A1048576")]
    [InlineData("R[2]C", 1048575, 1, "A1")]
    public void ToA1_loops_for_out_of_bounds_reference(string r1c1, int row, int col, string a1)
    {
        // In GUI, Excel loops over, if user enters out-of-bounds reference to a formula.
        Assert.Equal(a1, FormulaConverter.ToA1(r1c1, row, col));
    }
}