namespace ClosedXML.Parser.Tests.Lexers;

[TestClass]
public class IntraTableReferenceTokenTests
{
    [TestMethod]
    [DynamicData(nameof(Data))]
    public void Token_data_are_extracted_and_unescaped(string tokenText, StructuredReferenceSpecific expectedSpecifier, string expectedFirstColumn, string expectedLastColumn)
    {
        AssertFormula.AssertTokenType(tokenText, FormulaLexer.INTRA_TABLE_REFERENCE);
        TokenParser.ParseIntraTableReference(tokenText, out var specifier, out var firstColumn, out var lastColumn);

        Assert.AreEqual(expectedSpecifier, specifier);
        Assert.AreEqual(expectedFirstColumn, firstColumn);
        Assert.AreEqual(expectedLastColumn, lastColumn);

    }

    public static IEnumerable<object?[]> Data
    {
        get
        {
            // Portions area
            yield return new object?[] { "[#All]", StructuredReferenceSpecific.All, null, null };
            yield return new object?[] { "[#Data]", StructuredReferenceSpecific.Data, null, null };
            yield return new object?[] { "[#Headers]", StructuredReferenceSpecific.Headers, null, null };
            yield return new object?[] { "[#Totals]", StructuredReferenceSpecific.Totals, null, null };
            yield return new object?[] { "[#This Row]", StructuredReferenceSpecific.ThisRow, null, null };

            // Empty simple column
            yield return new object?[] { "[]", StructuredReferenceSpecific.None, null, null };

            // Simple column
            yield return new object?[] { "[Col]", StructuredReferenceSpecific.None, "Col", null };
            yield return new object?[] { "[Name with space]", StructuredReferenceSpecific.None, "Name with space", null };

            // Escaped characters
            yield return new object?[] { "['[']]", StructuredReferenceSpecific.None, "[]", null };
            yield return new object?[] { "['''#]", StructuredReferenceSpecific.None, "'#", null };
            yield return new object?[] { "['[']'''#]", StructuredReferenceSpecific.None, "[]'#", null };

            yield return new object?[] { "[[First]]", StructuredReferenceSpecific.None, "First", null };
            yield return new object?[] { "[[First]:[Last]]", StructuredReferenceSpecific.None, "First", "Last" };
            yield return new object?[] { "[[First]:Last]", StructuredReferenceSpecific.None, "First", "Last" };
            yield return new object?[] { "[First:[Last]]", StructuredReferenceSpecific.None, "First", "Last" };
            yield return new object?[] { "[First:Last]", StructuredReferenceSpecific.None, "First", "Last" };

            yield return new object?[] { "[[#All],[First]]", StructuredReferenceSpecific.All, "First", null };
            yield return new object?[] { "[[#Data],[First]:[Last]]", StructuredReferenceSpecific.Data, "First", "Last" };
            yield return new object?[] { "[[#Headers],[First]:Last]", StructuredReferenceSpecific.Headers, "First", "Last" };
            yield return new object?[] { "[[#Totals],First:[Last]]", StructuredReferenceSpecific.Totals, "First", "Last" };
            yield return new object?[] { "[[#This Row],First:Last]", StructuredReferenceSpecific.ThisRow, "First", "Last" };

            yield return new object?[] { "[[#Headers],[#Data],[Col]]", StructuredReferenceSpecific.Headers | StructuredReferenceSpecific.Data, "Col", null };
            yield return new object?[] { "[[#Headers],[#Data],[First col]:[Last col]]", StructuredReferenceSpecific.Headers | StructuredReferenceSpecific.Data, "First col", "Last col"};
            yield return new object?[] { "[[#Headers],[#Data],First:Last]", StructuredReferenceSpecific.Headers | StructuredReferenceSpecific.Data, "First", "Last" };
            yield return new object?[] { "[[#Headers],[#Data],[First]:Last]", StructuredReferenceSpecific.Headers | StructuredReferenceSpecific.Data, "First", "Last" };
            yield return new object?[] { "[[#Headers],[#Data],First:[Last]]", StructuredReferenceSpecific.Headers | StructuredReferenceSpecific.Data, "First", "Last" };

            // spaces are ignored
            yield return new object?[] { "[  [#Headers]  ,  [#Data]  ,  [First col]:[Last col]  ]", StructuredReferenceSpecific.Headers | StructuredReferenceSpecific.Data, "First col", "Last col" };
        }
    }
}