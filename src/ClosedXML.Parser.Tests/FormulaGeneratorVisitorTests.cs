namespace ClosedXML.Parser.Tests;

public class FormulaGeneratorVisitorTests
{
    [Theory]
    [InlineData("Old!B7:$D$10", "Old", "New", "New!B7:$D$10")]
    [InlineData("Old!B7:$D$10", "Old", "New sheet", "'New sheet'!B7:$D$10")]
    [InlineData("'Old Mike''s sheet'!B7:$D$10", "Old Mike's sheet", "New Mike's sheet", "'New Mike''s sheet'!B7:$D$10")]
    public void ModifySheet_can_rename_sheet_name(string formula, string oldSheetName, string newSheetName, string modifiedFormula)
    {
        var factory = new FormulaFactory { SheetMap = { { oldSheetName, newSheetName } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }

    private static void AssertChangesA1(string formula, FormulaGeneratorVisitor factory, string expected)
    {
        var ctx = new TransformContext(formula, 1, 1, isA1: true);
        var modifiedFormula = FormulaParser<TransformedSymbol, TransformedSymbol, TransformContext>.CellFormulaA1(formula, ctx, factory);
        Assert.Equal(expected, modifiedFormula.ToString(string.Empty.AsSpan()));
    }

    private class FormulaFactory : FormulaGeneratorVisitor
    {
        public Dictionary<string, string> SheetMap { get; } = new();

        protected override string ModifySheet(TransformContext ctx, string sheetName)
        {
            if (SheetMap.TryGetValue(sheetName, out var replacement))
                return replacement;

            return sheetName;
        }
    }
}