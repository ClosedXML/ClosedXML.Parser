namespace ClosedXML.Parser.Tests.Lexers;

[TestClass]
public class SingleSheetPrefixTokenTests
{
    [TestMethod]
    [DynamicData(nameof(Data))]
    public void Token_data_are_extracted_and_unescaped(string tokenText, int? expectedWorkbookIndex, string expectedSheetName)
    {
        AssertFormula.AssertTokenType(tokenText, FormulaLexer.SINGLE_SHEET_PREFIX);
        TokenParser.ParseSingleSheetPrefix(tokenText, out var workbookIndex, out var sheetName);

        Assert.AreEqual(expectedWorkbookIndex, workbookIndex);
        Assert.AreEqual(expectedSheetName, sheetName);
    }

    public static IEnumerable<object?[]> Data
    {
        get
        {
            yield return new object?[] { "sheet!", null, "sheet" };
            yield return new object?[] { "[7]sheet!", 7, "sheet" };
            yield return new object?[] { "'sheet name'!", null, "sheet name" };
            yield return new object?[] { "'[2]Monty''s'!", 2, "Monty's" };
            yield return new object?[] { "'[25]a''''''b'!", 25, "a'''b" };
        }
    }
}