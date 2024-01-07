namespace ClosedXML.Parser.Tests;

public class FormulaConverterToR1C1Tests
{
    [Theory]
    [InlineData("true", "true")]
    [InlineData("FALSE", "FALSE")]
    [InlineData("1", "1")]
    [InlineData("\"Text\"", "\"Text\"")]
    [InlineData("\"\"", "\"\"")]
    [InlineData("#DIV/0!", "#DIV/0!")]
    public void Constants(string a1, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, 1, 1));
    }

    [Theory]
    [InlineData(" { 1 } ", "{1}")]
    [InlineData("{1,2} ", "{1,2}")]
    [InlineData("{1;2} ", "{1;2}")]
    [InlineData("{1,2;3,4} ", "{1,2;3,4}")]
    [InlineData("{TRUE} ", "{TRUE}")]
    [InlineData("{#DIV/0!} ", "{#DIV/0!}")]
    [InlineData("{\"\"} ", "{\"\"}")]
    [InlineData("{\"Hello world\"} ", "{\"Hello world\"}")]
    public void Array(string a1, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, 1, 1));
    }

    [Theory]
    [InlineData("B4", 4, 2, "RC")] // A1 Relative
    [InlineData("B4", 5, 2, "R[-1]C")]
    [InlineData("B4", 4, 1, "RC[1]")]
    [InlineData("B4", 2, 1, "R[2]C[1]")]
    [InlineData("$B4", 4, 2, "RC2")] // A1 Mixed
    [InlineData("B$4", 4, 2, "R4C")]
    [InlineData("$B$4", 9, 7, "R4C2")] // A1 Absolute
    [InlineData("B4:B4", 2, 1, "R[2]C[1]")] // Both are same
    [InlineData("B4:B$4", 2, 1, "R[2]C[1]:R4C[1]")]
    [InlineData("C5:Z14", 2, 6, "R[3]C[-3]:R[12]C[20]")]
    public void Reference(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("'Sheet name'!$D6", 4, 1, "'Sheet name'!R[2]C4")]
    [InlineData("January!$D2", 4, 1, "January!R[-2]C4")]
    [InlineData("'A''B'!$D2", 4, 1, "'A''B'!R[-2]C4")]
    public void SheetReference(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("January:December!$D6", 4, 1, "January:December!R[2]C4")]
    [InlineData("'Johnny''s:Denny''s'!$D6", 4, 1, "'Johnny''s:Denny''s'!R[2]C4")]
    public void Reference3D(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[74]Sheet5!$D6", 4, 1, "[74]Sheet5!R[2]C4")]
    [InlineData("'[6]Johnny''s house'!$D6", 4, 1, "'[6]Johnny''s house'!R[2]C4")]
    public void ExternalSheetReference(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[74]Sheet1:Sheet7!$D6", 4, 1, "[74]Sheet1:Sheet7!R[2]C4")]
    [InlineData("'[6]Johnny''s:Danny''s'!$D6", 4, 1, "'[6]Johnny''s:Danny''s'!R[2]C4")]
    public void ExternalReference3D(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("RAND()", 4, 1, "RAND()")]
    [InlineData("SIN(F8)", 2, 3, "SIN(R[6]C[3])")]
    [InlineData("MOD(F8, $A$1)", 2, 3, "MOD(R[6]C[3],R1C1)")]
    [InlineData("IF(TRUE,,)", 2, 3, "IF(TRUE,,)")]
    public void Function(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("Sheet1!UDF.SHEET.FUNC($F$1)", 4, 1, "Sheet1!UDF.SHEET.FUNC(R1C6)")]
    [InlineData("'Johnny''s'!UDF.SHEET.FUNC($F$1)", 4, 1, "'Johnny''s'!UDF.SHEET.FUNC(R1C6)")]
    public void SheetFunction(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[4]!UDF.SHEET.FUNC($F$1)", 4, 1, "[4]!UDF.SHEET.FUNC(R1C6)")]
    public void ExternalFunction(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[4]Sheet1!UDF.SHEET.FUNC($F$1)", 4, 1, "[4]Sheet1!UDF.SHEET.FUNC(R1C6)")]
    [InlineData("'[7]Johnny''s'!UDF.SHEET.FUNC($F$1)", 4, 1, "'[7]Johnny''s'!UDF.SHEET.FUNC(R1C6)")]
    public void ExternalSheetFunction(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("$A$5(TRUE, 7)", 10, 15, "R5C1(TRUE,7)")]
    public void CellFunction(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[]", 10, 15, "[]")]
    [InlineData("[#Headers]", 10, 15, "[#Headers]")]
    [InlineData("[#Data]", 10, 15, "[#Data]")]
    [InlineData("[#Totals]", 10, 15, "[#Totals]")]
    [InlineData("[#All]", 10, 15, "[#All]")]
    [InlineData("[#This Row]", 10, 15, "[#This Row]")]
    [InlineData("[[#Headers],[#Data]]", 10, 15, "[[#Headers],[#Data]]", Skip = "Parser fail")]
    [InlineData("[[#Data],[#Totals]]", 10, 15, "[[#Data],[#Totals]]", Skip = "Parser fail")]
    [InlineData("[Column]", 10, 15, "[Column]")]
    [InlineData("[Space column]", 10, 15, "[Space column]")]
    [InlineData("[[#Data],[Column]]", 10, 15, "[[#Data],[Column]]")]
    [InlineData("[[#Data],[Space column]]", 10, 15, "[[#Data],[Space column]]")]
    [InlineData("[[#Headers],[#Data],[Column]]", 10, 15, "[[#Headers],[#Data],[Column]]")]
    [InlineData("[[#Data],[#Totals],[Space column]]", 10, 15, "[[#Data],[#Totals],[Space column]]")]
    [InlineData("[[#All],[Column 1]:[Column 2]]", 10, 15, "[[#All],[Column 1]:[Column 2]]")]
    [InlineData("[[#Data],[#Totals],[Column 1]:[Column 2]]", 10, 15, "[[#Data],[#Totals],[Column 1]:[Column 2]]")]
    public void StructureReference(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("Table1[[#Data],[#Totals],[Column 1]:[Column 2]]", 10, 15, "Table1[[#Data],[#Totals],[Column 1]:[Column 2]]")]
    public void TableStructureReference(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[1]Table1[Column]", 10, 15, "[1]Table1[Column]", Skip = "Parser fail")]
    public void ExternalStructureReference(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData(" some_name + other_name", 1, 1, "some_name+other_name")]
    public void Name(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("Sheet1!defined_name", 1, 1, "Sheet1!defined_name")]
    [InlineData("'Sheet name'!defined_name", 1, 1, "'Sheet name'!defined_name")]
    public void SheetName(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[4]!defined_name", 1, 1, "[4]!defined_name")]
    public void ExternalName(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("[0]Sheet5!name", 1, 1, "[0]Sheet5!name")]
    [InlineData("'[4]Happy sheet'!data", 1, 1, "'[4]Happy sheet'!data")]
    public void ExternalSheetName(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("+B3", 1, 1, "+R[2]C[1]")]
    [InlineData("-8", 1, 1, "-8")]
    [InlineData("100 %", 1, 1, "100%")]
    [InlineData("@D8", 2, 5, "@R[6]C[-1]")]
    [InlineData("D8#", 2, 5, "R[6]C[-1]#")]
    public void Unary(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

    [Theory]
    [InlineData("(1 + 2) / 4", 1, 1, "(1+2)/4")]
    [InlineData("(1+((3 + A4)))", 1, 1, "(1+((3+R[3]C)))")]
    public void Nested(string a1, int row, int col, string r1c1)
    {
        Assert.Equal(r1c1, FormulaConverter.ToR1C1(a1, row, col));
    }

}