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
}