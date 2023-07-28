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

    [Theory]
    [MemberData(nameof(DisplayStringR1C1))]
    public void DisplayStringR1C1_displays_reference_in_R1C1_style(Reference reference, string expectedString)
    {
        Assert.Equal(expectedString, reference.GetDisplayStringR1C1());
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            yield return new object[] { new Reference(Relative, 1, Relative, 1), "A1" };
            yield return new object[] { new Reference(Absolute, 28, Relative, 14), "$AB14" };
            yield return new object[] { new Reference(Relative, 26, Absolute, 4), "Z$4" };
            yield return new object[] { new Reference(Absolute, 3, Absolute, 264), "$C$264" };
        }
    }

    public static IEnumerable<object[]> DisplayStringR1C1
    {
        get
        {
            yield return new object[] { new Reference(Relative, 1, Relative, 1), "R[1]C[1]" };
            yield return new object[] { new Reference(None, 0, Relative, 105), "R[105]" };
            yield return new object[] { new Reference(Relative, -7, None, 0), "C[-7]" };
            yield return new object[] { new Reference(Absolute, 1, Absolute, 1), "R1C1" };
            yield return new object[] { new Reference(Absolute, 8, None, 0), "C8" };
            yield return new object[] { new Reference(None, 0, Absolute, 1), "R1" };
            yield return new object[] { new Reference(Relative, 0, None, 0), "C" };
            yield return new object[] { new Reference(None, 0, Relative, 0), "R" };
        }
    }
}