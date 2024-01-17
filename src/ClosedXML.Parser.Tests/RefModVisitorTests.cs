namespace ClosedXML.Parser.Tests;

public class RefModVisitorTests
{
    [Theory]
    [InlineData("Old!B7:$D$10", "Old", "New", "New!B7:$D$10")]
    [InlineData("Old!B7:$D$10", "Old", "New sheet", "'New sheet'!B7:$D$10")]
    [InlineData("'Old Mike''s sheet'!B7:$D$10", "Old Mike's sheet", "New Mike's sheet", "'New Mike''s sheet'!B7:$D$10")]
    public void ModifySheet_can_rename_sheet_name(string formula, string oldSheetName, string newSheetName, string modifiedFormula)
    {
        var factory = new FormulaVisitor { SheetMap = { { oldSheetName, newSheetName } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Theory]
    [InlineData("Old!#REF!", "Old", null, "#REF!#REF!")]
    [InlineData("Old!#REF!", "Old", "New", "New!#REF!")]
    public void ErrorNode_can_modify_sheet(string formula, string oldSheetName, string? newSheetName, string modifiedFormula)
    {
        var factory = new FormulaVisitor { SheetMap = { { oldSheetName, newSheetName } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Theory]
    [InlineData("Old!B$5", "Old", null, "#REF!")]
    [InlineData("Old!B:D", "Old", "Shiny", "Shiny!B:D")]
    public void SheetReference_can_modify_sheet(string formula, string oldSheetName, string? newSheetName, string modifiedFormula)
    {
        var factory = new FormulaVisitor { SheetMap = { { oldSheetName, newSheetName } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Theory]
    [InlineData("Old!F(5)", "Old", null, "#REF!")]
    [InlineData("Old!F(7)", "Old", "Shiny", "Shiny!F(7)")]
    public void SheetFunction_can_modify_sheet(string formula, string oldSheetName, string? newSheetName, string modifiedFormula)
    {
        var factory = new FormulaVisitor { SheetMap = { { oldSheetName, newSheetName } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Theory]
    [InlineData("Sheet1:Sheet5!A1", "Sheet1", null, "#REF!")]
    [InlineData("Sheet1:Sheet5!A1", "Sheet1", "New sheet", "'New sheet:Sheet5'!A1")]
    [InlineData("Sheet1:Sheet5!A1", "Sheet5", "Sheet9", "Sheet1:Sheet9!A1")]
    public void Reference3D_can_modify_sheet(string formula, string oldSheetName, string? newSheetName, string modifiedFormula)
    {
        var factory = new FormulaVisitor { SheetMap = { { oldSheetName, newSheetName } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }
    
    [Theory]
    [InlineData("Sheet!Name", "Sheet", null, "#REF!")]
    [InlineData("Sheet!Name", "Sheet", "New Sheet", "'New Sheet'!Name")]
    public void SheetName_can_modify_sheet(string formula, string oldSheetName, string? newSheetName, string modifiedFormula)
    {
        var factory = new FormulaVisitor { SheetMap = { { oldSheetName, newSheetName } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }
    
    private static void AssertChangesA1(string formula, RefModVisitor visitor, string expected)
    {
        var ctx = new ModContext(formula, 1, 1, isA1: true);
        var modifiedFormula = FormulaParser<TransformedSymbol, TransformedSymbol, ModContext>.CellFormulaA1(formula, ctx, visitor);
        Assert.Equal(expected, modifiedFormula.ToString(string.Empty.AsSpan()));
    }

    private class FormulaVisitor : RefModVisitor
    {
        public Dictionary<string, string?> SheetMap { get; } = new();

        protected override string? ModifySheet(ModContext ctx, string sheetName)
        {
            if (SheetMap.TryGetValue(sheetName, out var replacement))
                return replacement;

            return sheetName;
        }
    }
}
