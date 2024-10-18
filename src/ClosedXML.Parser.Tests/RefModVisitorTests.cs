namespace ClosedXML.Parser.Tests;

public class RefModVisitorTests
{
    #region ModifySheet

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

    #endregion

    [Theory]
    [InlineData("5 + !$B1", "$B1", "$7:$9", "5 + !$7:$9")]
    [InlineData("5 + !$B1", "$B1", null, "5 + !#REF!")]
    public void Bang_references_is_modified(string formula, string reference, string? replacement, string modifiedFormula)
    {
        var factory = new ShiftReferenceVisitor { ReferenceMap = { { reference, replacement } } };
        AssertChangesA1(formula, factory, modifiedFormula);
    }

    private static void AssertChangesA1(string formula, RefModVisitor visitor, string expected)
    {
        var ctx = new ModContext(formula, "Sheet", 1, 1, isA1: true);
        var modifiedFormula = FormulaParser<TransformedSymbol, TransformedSymbol, ModContext>.CellFormulaA1(formula, ctx, visitor);
        Assert.Equal(expected, modifiedFormula.ToString(string.Empty.AsSpan()));
    }

    private class FormulaVisitor : RefModVisitor
    {
        public Dictionary<string, string?> SheetMap { get; } = new();

        protected override string? ModifySheet(ModContext ctx, string sheetName)
        {
            return SheetMap.GetValueOrDefault(sheetName, sheetName);
        }
    }

    private class ShiftReferenceVisitor : RefModVisitor
    {
        public Dictionary<string, string?> ReferenceMap { get; } = new();

        internal override ReferenceArea? ModifyRef(ModContext ctx, ReferenceArea reference)
        {
            if (ReferenceMap.TryGetValue(reference.GetDisplayStringA1(), out var replacement))
                return replacement is not null ? ReferenceParser.ParseA1(replacement) : null;

            return reference;
        }
    }
}
