using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests;

public class ReferenceTests
{
    [Theory]
    [MemberData(nameof(DisplayStringA1))]
    public void DisplayStringA1_displays_reference_in_A1_style(Reference reference, string expectedString)
    {
        Assert.Equal(expectedString, reference.GetDisplayStringA1());
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            yield return new object[]
            {
                new Reference(Relative, 1, Relative, 1),
                "A1"
            };

            yield return new object[]
            {
                new Reference(Absolute, 28, Relative, 14),
                "$AB14"
            };

            yield return new object[]
            {
                new Reference(Relative, 26, Absolute, 4),
                "Z$4"
            };


            yield return new object[]
            {
                new Reference(Absolute, 3, Absolute, 264),
                "$C$264"
            };
        }
    }
}