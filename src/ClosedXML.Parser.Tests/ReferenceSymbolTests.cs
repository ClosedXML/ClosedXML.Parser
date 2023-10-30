namespace ClosedXML.Parser.Tests;

public class ReferenceSymbolTests
{
    [Theory]
    [MemberData(nameof(DisplayStringA1))]
    public void DisplayStringA1_displays_reference_in_A1_style(ReferenceSymbol reference, string expectedString)
    {
        Assert.Equal(expectedString, reference.GetDisplayStringA1());
    }

    [Theory]
    [MemberData(nameof(DisplayStringR1C1))]
    public void DisplayStringR1C1_displays_reference_in_R1C1_style(ReferenceSymbol reference, string expectedString)
    {
        Assert.Equal(expectedString, reference.GetDisplayStringR1C1());
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1), new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1)),
                "A1"
            };

            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1), new RowCol(ReferenceAxisType.Relative, 5, ReferenceAxisType.Relative, 3, A1)),
                "A1:C5"
            };
        }
    }

    public static IEnumerable<object[]> DisplayStringR1C1
    {
        get
        {
            yield return new object[]
            {
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Absolute, 7, ReferenceAxisType.Absolute, 1, R1C1), new RowCol(ReferenceAxisType.Absolute, 7, ReferenceAxisType.Absolute, 1, R1C1)),
                "R7C1"
            };

            yield return new object[]
            {
                new ReferenceSymbol(new RowCol(ReferenceAxisType.Absolute, 7, ReferenceAxisType.Absolute, 1, R1C1), new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0, R1C1)),
                "R7C1:R"
            };
       }
    }
}