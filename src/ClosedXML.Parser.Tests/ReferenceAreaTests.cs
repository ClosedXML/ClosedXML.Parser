using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests;

public class ReferenceAreaTests
{
    [Theory]
    [MemberData(nameof(DisplayStringA1))]
    public void DisplayStringA1_displays_reference_in_A1_style(ReferenceArea reference, string expectedString)
    {
        Assert.Equal(expectedString, reference.GetDisplayStringA1());
        Assert.Equal(reference, TokenParser.ParseReference(expectedString, true));
    }

    [Theory]
    [MemberData(nameof(DisplayStringR1C1))]
    public void DisplayStringR1C1_displays_reference_in_R1C1_style(ReferenceArea reference, string expectedString)
    {
        Assert.Equal(expectedString, reference.GetDisplayStringR1C1());
        Assert.Equal(reference, TokenParser.ParseReference(expectedString, false));
    }

    [Theory]
    [MemberData(nameof(ParseA1TestCases))]
    public void ParseA1_parses_cell_area_or_rowspan_or_colspan(string text, ReferenceArea expectedReference)
    {
        Assert.Equal(expectedReference, ReferenceArea.ParseA1(text));
    }

    [Fact]
    public void ParseA1_requires_argument()
    {
        Assert.Throws<ArgumentNullException>(() => ReferenceArea.ParseA1(null!));
    }

    [Fact]
    public void ParseA1_throws_on_non_references()
    {
        Assert.Throws<ParsingException>(() => ReferenceArea.ParseA1("HELLO"));
    }

    public static IEnumerable<object[]> ParseA1TestCases
    {
        get
        {
            yield return new object[]
            {
                "$C$2",
                new ReferenceArea(new RowCol(Absolute, 2, Absolute, 3, A1)),
            };
            yield return new object[]
            {
                "AB123",
                new ReferenceArea(new RowCol(Relative, 123, Relative, 28, A1)),
            };
            yield return new object[]
            {
                "$C$2:E7",
                new ReferenceArea(new RowCol(Absolute, 2, Absolute, 3, A1), new RowCol(Relative, 7, Relative, 5, A1)),
            };
            yield return new object[]
            {
                "$C:F",
                new ReferenceArea(new RowCol(None, 0, Absolute, 3, A1), new RowCol(None, 0, Relative, 6, A1)),
            };
            yield return new object[]
            {
                "10:$15",
                new ReferenceArea(new RowCol(Relative, 10, None, 0, A1), new RowCol(Absolute, 15, None, 0, A1)),
            };
        }
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 1, Relative, 1, A1), new RowCol(Relative, 1, Relative, 1, A1)),
                "A1"
            };

            // When both corners are same, only one is rendered.
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 1, Relative, 1, A1), new RowCol(Relative, 5, Relative, 3, A1)),
                "A1:C5"
            };

            // Row span
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 6, None, 0, A1), new RowCol(Relative, 6, None, 0, A1)),
                "6:6"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 6, None, 0, A1), new RowCol(Relative, 8, None, 0, A1)),
                "6:8"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Absolute, 65, None, 0, A1), new RowCol(Absolute, 745, None, 0, A1)),
                "$65:$745"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 79, None, 0, A1), new RowCol(Absolute, 999, None, 0, A1)),
                "79:$999"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Absolute, 79, None, 0, A1), new RowCol(Relative, 999, None, 0, A1)),
                "$79:999"
            };

            // Col span
            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 5, A1), new RowCol(None, 0, Relative, 5, A1)),
                "E:E"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 2, A1), new RowCol(None, 0, Relative, 4, A1)),
                "B:D"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Absolute, 27, A1), new RowCol(None, 0, Absolute, 53, A1)),
                "$AA:$BA"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 96, A1), new RowCol(None, 0, Absolute, 6663, A1)),
                "CR:$IVG"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Absolute, 96, A1), new RowCol(None, 0, Relative, 6663, A1)),
                "$CR:IVG"
            };
        }
    }

    public static IEnumerable<object[]> DisplayStringR1C1
    {
        get
        {
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Absolute, 7, Absolute, 1, R1C1), new RowCol(Absolute, 7, Absolute, 1, R1C1)),
                "R7C1"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Absolute, 7, Absolute, 1, R1C1), new RowCol(Relative, 0, None, 0, R1C1)),
                "R7C1:R"
            };

            // Row span
            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 6, None, 0, R1C1), new RowCol(Relative, 6, None, 0, R1C1)),
                "R[6]"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 6, None, 0, R1C1), new RowCol(Relative, 8, None, 0, R1C1)),
                "R[6]:R[8]"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Absolute, 65, None, 0, R1C1), new RowCol(Absolute, 745, None, 0, R1C1)),
                "R65:R745"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Relative, 79, None, 0, R1C1), new RowCol(Absolute, 999, None, 0, R1C1)),
                "R[79]:R999"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(Absolute, 79, None, 0, R1C1), new RowCol(Relative, 999, None, 0, R1C1)),
                "R79:R[999]"
            };

            // Col span
            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 2, R1C1), new RowCol(None, 0, Relative, 2, R1C1)),
                "C[2]"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 2, R1C1), new RowCol(None, 0, Relative, 4, R1C1)),
                "C[2]:C[4]"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Absolute, 27, R1C1), new RowCol(None, 0, Absolute, 53, R1C1)),
                "C27:C53"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Relative, 96, R1C1), new RowCol(None, 0, Absolute, 663, R1C1)),
                "C[96]:C663"
            };

            yield return new object[]
            {
                new ReferenceArea(new RowCol(None, 0, Absolute, 96, R1C1), new RowCol(None, 0, Relative, 663, R1C1)),
                "C96:C[663]"
            };
        }
    }
}