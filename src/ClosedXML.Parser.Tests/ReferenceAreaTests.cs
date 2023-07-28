namespace ClosedXML.Parser.Tests;

public class ReferenceAreaTests
{
    [Theory]
    [MemberData(nameof(DisplayStringA1))]
    public void DisplayStringA1_displays_reference_in_A1_style(ReferenceArea area, string expectedString)
    {
        Assert.Equal(expectedString, area.GetDisplayStringA1());
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceArea(new Reference(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1), new Reference(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1)),
                "A1"
            };

            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceArea(new Reference(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1), new Reference(ReferenceAxisType.Relative, 3, ReferenceAxisType.Relative, 5)),
                "A1:C5"
            };
        }
    }
}