namespace ClosedXML.Parser.Tests;

public class ReferenceAreaTests
{
    [Theory]
    [MemberData(nameof(DisplayStringA1))]
    public void DisplayStringA1_displays_reference_in_A1_style(ReferenceArea area, string expectedString)
    {
        Assert.Equal(expectedString, area.GetDisplayStringA1());
    }

    [Theory]
    [MemberData(nameof(DisplayStringR1C1))]
    public void DisplayStringR1C1_displays_reference_in_R1C1_style(ReferenceArea area, string expectedString)
    {
        Assert.Equal(expectedString, area.GetDisplayStringR1C1());
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceArea(new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1), new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1)),
                "A1"
            };

            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceArea(new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1), new RowCol(ReferenceAxisType.Relative, 3, ReferenceAxisType.Relative, 5)),
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
                new ReferenceArea(new RowCol(ReferenceAxisType.Absolute, 1, ReferenceAxisType.Absolute, 7), new RowCol(ReferenceAxisType.Absolute, 1, ReferenceAxisType.Absolute, 7)),
                "R7C1"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(ReferenceAxisType.Absolute, 1, ReferenceAxisType.Absolute, 7), new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0)),
                "R7C1:R"
            };
       }
    }
}