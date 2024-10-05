using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser.Tests;

public class ReferenceParserTests
{
    [Theory]
    [MemberData(nameof(ParseA1TestCases))]
    public void ParseA1_parses_cell_area_or_rowspan_or_colspan(string text, ReferenceArea expectedReference)
    {
        Assert.Equal(expectedReference, ReferenceParser.ParseA1(text));
    }

    [Fact]
    public void ParseA1_requires_argument()
    {
        Assert.Throws<ArgumentNullException>(() => ReferenceParser.ParseA1(null!));
    }

    [Fact]
    public void ParseA1_throws_on_non_references()
    {
        Assert.Throws<ParsingException>(() => ReferenceParser.ParseA1("HELLO"));
    }

    [Theory]
    [MemberData(nameof(ParseA1TestCases))]
    public void TryParseA1_parses_cell_area_or_rowspan_or_colspan(string text, ReferenceArea expectedReference)
    {
        var success = ReferenceParser.TryParseA1(text, out var area);
        Assert.True(success);
        Assert.Equal(expectedReference, area);
    }

    [Fact]
    public void TryParseA1_requires_argument()
    {
        Assert.Throws<ArgumentNullException>(() => ReferenceParser.TryParseA1(null!, out _));
    }

    [Fact]
    public void ParseA1_returns_false_on_non_references()
    {
        var success = ReferenceParser.TryParseA1("HELLO", out var area);
        Assert.False(success);
        Assert.Equal(default, area);
    }

    [Theory]
    [MemberData(nameof(ParseSheetA1TestCases))]
    public void TryParseSheetA1_accepts_area_or_rowspan_or_colspan_with_sheet(string text, string expectedSheet, ReferenceArea expectedArea)
    {
        var success = ReferenceParser.TryParseSheetA1(text, out var sheet, out var area);
        Assert.True(success);
        Assert.Equal(expectedSheet, sheet);
        Assert.Equal(expectedArea, area);
    }

    [Fact]
    public void TryParseSheetA1_cant_parse_workbook_index()
    {
        var success = ReferenceParser.TryParseSheetA1("[1]Sheet!A1", out _, out _);
        Assert.False(success);
    }

    [Fact]
    public void TryParseSheetA1_cant_parse_reference_without_sheet()
    {
        var success = ReferenceParser.TryParseSheetA1("A1", out _, out _);
        Assert.False(success);
    }

    [Fact]
    public void TryParseSheetA1_requires_argument()
    {
        Assert.Throws<ArgumentNullException>(() => ReferenceParser.TryParseSheetA1(null!, out _, out _));
    }

    [Theory]
    [InlineData("Sheet!Name", "Sheet", "Name")]
    [InlineData("'Hello World'!Name", "Hello World", "Name")]
    [InlineData("' John''s World! '!Name", " John's World! ", "Name")]
    public void TryParseSheetName_parses_sheet_and_name(string text, string expectedSheet, string expectedName)
    {
        var success = ReferenceParser.TryParseSheetName(text, out var sheet, out var name);

        Assert.True(success);
        Assert.Equal(expectedSheet, sheet);
        Assert.Equal(expectedName, name);
    }

    [Fact]
    public void TryParseSheetName_requires_text()
    {
        Assert.Throws<ArgumentNullException>(() => ReferenceParser.TryParseSheetName(null!, out _, out _));
    }

    [Theory]
    [InlineData("Name")]
    [InlineData("some_name")]
    [InlineData("A1")]
    public void TryParseSheetName_cant_parse_pure_name(string text)
    {
        var success = ReferenceParser.TryParseSheetName(text, out _, out _);
        Assert.False(success);
    }

    [Theory]
    [MemberData(nameof(ParseA1TestCases))]
    public void TryParseA1_unified_can_parse_local_reference(string text, ReferenceArea expectedReference)
    {
        var success = ReferenceParser.TryParseA1(text, out var sheet, out var reference);
        Assert.True(success);
        Assert.Null(sheet);
        Assert.Equal(expectedReference, reference);
    }

    [Theory]
    [MemberData(nameof(ParseSheetA1TestCases))]
    public void TryParseA1_unified_can_parse_sheet_reference(string text, string expectedSheet, ReferenceArea expectedReference)
    {
        var success = ReferenceParser.TryParseA1(text, out var sheet, out var reference);
        Assert.True(success);
        Assert.Equal(expectedSheet, sheet);
        Assert.Equal(expectedReference, reference);
    }

    [Theory]
    [InlineData("Sheet!Name")]
    [InlineData("Name")]
    [InlineData("1")]
    public void TryParseA1_unified_cant_parse_anything_but_reference(string text)
    {
        var success = ReferenceParser.TryParseA1(text, out _, out _);
        Assert.False(success);
    }

    [Fact]
    public void TryParseA1_unified_requires_argument()
    {
        Assert.Throws<ArgumentNullException>(() => ReferenceParser.TryParseA1(null!, out _, out _));
    }

    public static IEnumerable<object[]> ParseSheetA1TestCases
    {
        get
        {
            yield return new object[]
            {
                "Sheet!$C$2",
                "Sheet",
                new ReferenceArea(new RowCol(Absolute, 2, Absolute, 3, A1)),
            };
            yield return new object[]
            {
                "' ''John''s'' Shop! '!C2",
                " 'John's' Shop! ",
                new ReferenceArea(new RowCol(Relative, 2, Relative, 3, A1)),
            };
            yield return new object[]
            {
                "Sheet!A1:B2",
                "Sheet",
                new ReferenceArea(new RowCol(Relative, 1, Relative, 1, A1), new RowCol(Relative, 2, Relative, 2, A1)),
            };
            yield return new object[]
            {
                "'Some Sheet'!C:D",
                "Some Sheet",
                new ReferenceArea(new RowCol(None, 0, Relative, 3, A1), new RowCol(None, 0, Relative, 4, A1)),
            };
            yield return new object[]
            {
                "'!!WARN'!10:$15",
                "!!WARN",
                new ReferenceArea(new RowCol(Relative, 10, None, 0, A1), new RowCol(Absolute, 15, None, 0, A1)),
            };
        }
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
}
