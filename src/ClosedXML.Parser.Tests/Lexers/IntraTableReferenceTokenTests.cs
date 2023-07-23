using Xunit;

namespace ClosedXML.Parser.Tests.Lexers;

public class IntraTableReferenceTokenTests
{
    [Theory]
    [MemberData(nameof(Data))]
    public void Token_data_are_extracted_and_unescaped(string tokenText, StructuredReferenceArea expectedArea, string expectedFirstColumn, string expectedLastColumn)
    {
        AssertFormula.AssertTokenType(tokenText, FormulaLexer.INTRA_TABLE_REFERENCE);
        TokenParser.ParseIntraTableReference(tokenText, out var area, out var firstColumn, out var lastColumn);

        Assert.Equal(expectedArea, area);
        Assert.Equal(expectedFirstColumn, firstColumn);
        Assert.Equal(expectedLastColumn, lastColumn);
    }

    public static IEnumerable<object?[]> Data
    {
        get
        {
            // Portions area
            // INTRA_TABLE_REFERENCE : KEYWORD
            yield return new object?[] { "[#All]", StructuredReferenceArea.All, null, null };
            yield return new object?[] { "[#Data]", StructuredReferenceArea.Data, null, null };
            yield return new object?[] { "[#Headers]", StructuredReferenceArea.Headers, null, null };
            yield return new object?[] { "[#Totals]", StructuredReferenceArea.Totals, null, null };
            yield return new object?[] { "[#This Row]", StructuredReferenceArea.ThisRow, null, null };

            // Empty simple column, per grammar, the SIMPLE_COLUMN_NAME is optional
            // INTRA_TABLE_REFERENCE : '[' SIMPLE_COLUMN_NAME? ']' 
            yield return new object?[] { "[]", StructuredReferenceArea.None, null, null };

            // Simple column
            // INTRA_TABLE_REFERENCE : '[' SIMPLE_COLUMN_NAME? ']' 
            yield return new object?[] { "[Col]", StructuredReferenceArea.None, "Col", null };
            yield return new object?[] { "[Name with space]", StructuredReferenceArea.None, "Name with space", null };

            // Escaped characters
            // INTRA_TABLE_REFERENCE : '[' SIMPLE_COLUMN_NAME? ']'
            // where column name is a possible value of a ESCAPE_COLUMN_CHARACTER
            yield return new object?[] { "['[']]", StructuredReferenceArea.None, "[]", null };
            yield return new object?[] { "['''#]", StructuredReferenceArea.None, "'#", null };
            yield return new object?[] { "['[']'''#]", StructuredReferenceArea.None, "[]'#", null };

            // INTRA_TABLE_REFERENCE : SPACED_LBRACKET INNER_REFERENCE SPACED_RBRACKET
            // where inner reference is `COLUMN_RANGE : COLUMN(':' COLUMN)?`
            yield return new object?[] { "[[First]]", StructuredReferenceArea.None, "First", null };
            yield return new object?[] { "[[First]:[Last]]", StructuredReferenceArea.None, "First", "Last" };
            yield return new object?[] { "[[First]:Last]", StructuredReferenceArea.None, "First", "Last" };
            yield return new object?[] { "[First:[Last]]", StructuredReferenceArea.None, "First", "Last" };
            yield return new object?[] { "[First:Last]", StructuredReferenceArea.None, "First", "Last" };

            // fragment INNER_REFERENCE : KEYWORD_LIST SPACED_COMMA COLUMN_RANGE
            // where KEYWORD_LIST is just a KEYWORD
            yield return new object?[] { "[[#All],[First]]", StructuredReferenceArea.All, "First", null };
            yield return new object?[] { "[[#Data],[First]:[Last]]", StructuredReferenceArea.Data, "First", "Last" };
            yield return new object?[] { "[[#Headers],[First]:Last]", StructuredReferenceArea.Headers, "First", "Last" };
            yield return new object?[] { "[[#Totals],First:[Last]]", StructuredReferenceArea.Totals, "First", "Last" };
            yield return new object?[] { "[[#This Row],First:Last]", StructuredReferenceArea.ThisRow, "First", "Last" };

            // fragment INNER_REFERENCE : KEYWORD_LIST SPACED_COMMA COLUMN_RANGE
            // where KEYWORD_LIST | '[#Headers]' SPACED_COMMA '[#Data]' | '[#Data]' SPACED_COMMA '[#Totals]'
            yield return new object?[] { "[[#Headers],[#Data],[Col]]", StructuredReferenceArea.Headers | StructuredReferenceArea.Data, "Col", null };
            yield return new object?[] { "[[#Headers],[#Data],[First col]:[Last col]]", StructuredReferenceArea.Headers | StructuredReferenceArea.Data, "First col", "Last col"};
            yield return new object?[] { "[[#Headers],[#Data],First:Last]", StructuredReferenceArea.Headers | StructuredReferenceArea.Data, "First", "Last" };
            yield return new object?[] { "[[#Headers],[#Data],[First]:Last]", StructuredReferenceArea.Headers | StructuredReferenceArea.Data, "First", "Last" };
            yield return new object?[] { "[[#Headers],[#Data],First:[Last]]", StructuredReferenceArea.Headers | StructuredReferenceArea.Data, "First", "Last" };

            // spaces are ignored
            yield return new object?[] { "[  [#Headers]  ,  [#Data]  ,  [First col]:[Last col]  ]", StructuredReferenceArea.Headers | StructuredReferenceArea.Data, "First col", "Last col" };
        }
    }
}