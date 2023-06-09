﻿namespace ClosedXML.Parser.Tests.Lexers;

// Tests of parsing SHEET_RANGE_PREFIX
[TestClass]
public class SheetRangePrefixTokenTests
{
    [TestMethod]
    [DynamicData(nameof(Data))]
    public void Token_data_are_extracted_and_unescaped(string tokenText, int? expectedWorkbookIndex, string expectedFirstSheetName, string expectedSecondSheetName)
    {
        AssertFormula.AssertTokenType(tokenText, FormulaLexer.SHEET_RANGE_PREFIX);
        TokenParser.ParseSheetRangePrefix(tokenText, out var workbookIndex, out var firstSheetName, out var secondSheetName);

        Assert.AreEqual(expectedWorkbookIndex, workbookIndex);
        Assert.AreEqual(expectedFirstSheetName, firstSheetName);
        Assert.AreEqual(expectedSecondSheetName, secondSheetName);
    }

    public static IEnumerable<object?[]> Data
    {
        get
        {
            yield return new object?[] { "[1]first:second!", 1, "first", "second" };
            yield return new object?[] { "first:second!", null, "first", "second" };

            // No escape, but enclosed in tick
            yield return new object?[] { "'[1]first:second'!", 1, "first", "second" };
            yield return new object?[] { "'first:second'!", null, "first", "second" };

            // Test correct escaping
            yield return new object?[] { "'Monty''s:Johnny''s'!", null, "Monty's", "Johnny's" };

            // multiple escapes
            yield return new object?[] { "'[7]a''''''b:c''''d'!", 7, "a'''b", "c''d" };

            // single character name
            yield return new object?[] { "'[6]a:b'!", 6, "a", "b" };
        }
    }
}