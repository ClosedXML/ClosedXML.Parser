using Antlr4.Runtime;

namespace ClosedXML.Parser.Tests;

[TestClass]
public class CellReferenceTests
{
    private const int MaxCol = 16384;
    private const int MaxRow = 1048576;

    [TestMethod]
    public void Parse_a1_cell()
    {
        Assert.AreEqual(new CellArea(new CellReference(true, 2, true, 3)), ParseCellReference("$B$3"));
        Assert.AreEqual(new CellArea(1, 1), ParseCellReference("A1"));
        Assert.AreEqual(new CellArea(16384, 1), ParseCellReference("XFD1"));
        Assert.AreEqual(new CellArea(1, 1048576), ParseCellReference("A1048576"));
        Assert.AreEqual(new CellArea(new CellReference(true, 16384, true, 1048576)), ParseCellReference("$XFD$1048576"));
    }

    [TestMethod]
    public void Parse_row_range()
    {
        Assert.AreEqual(new CellArea(new CellReference(true, 1, false, 1), new CellReference(true, MaxCol, false, 1)), ParseCellReference("1:1"));
        Assert.AreEqual(new CellArea(new CellReference(true, 1, true, 5), new CellReference(true, MaxCol, false, 10)), ParseCellReference("$5:10"));
        Assert.AreEqual(new CellArea(new CellReference(true, 1, false, 7), new CellReference(true, MaxCol, true, 3)), ParseCellReference("7:$3"));
        Assert.AreEqual(new CellArea(new CellReference(true, 1, true, 1048576), new CellReference(true, MaxCol, true, 1048576)), ParseCellReference("$1048576:$1048576"));
    }

    [TestMethod]
    public void Parse_column_range()
    {
        Assert.AreEqual(new CellArea(new CellReference(false, 1, true, 1), new CellReference(false, 1, true, MaxRow)), ParseCellReference("A:A"));
        Assert.AreEqual(new CellArea(new CellReference(false, 491, true, 1), new CellReference(false, 514, true, MaxRow)), ParseCellReference("RW:ST"));
        Assert.AreEqual(new CellArea(new CellReference(true, 3, true, 1), new CellReference(false, 4, true, MaxRow)), ParseCellReference("$C:D"));
        Assert.AreEqual(new CellArea(new CellReference(false, 5, true, 1), new CellReference(true, 3, true, MaxRow)), ParseCellReference("E:$C"));
        Assert.AreEqual(new CellArea(new CellReference(true, MaxCol, true, 1), new CellReference(true, MaxCol, true, MaxRow)), ParseCellReference("$XFD:$XFD"));
    }

    [TestMethod]
    public void Parse_area()
    {
        Assert.AreEqual(new CellArea(new CellReference(false, 1, false, 1)), ParseCellReference("A1:A1"));
        Assert.AreEqual(new CellArea(new CellReference(false, 26, false, 1), new CellReference(28, 25)), ParseCellReference("Z1:AB25"));
        Assert.AreEqual(new CellArea(new CellReference(true, MaxCol - 1, true, MaxRow - 1), new CellReference(true, MaxCol, true, MaxRow)), ParseCellReference("$XFC$1048575:$XFD$1048576"));
    }

    private static CellArea ParseCellReference(string formula)
    {
        var lexer = new FormulaLexer(new CodePointCharStream(formula), TextWriter.Null, TextWriter.Null);
        lexer.RemoveErrorListeners();
        var parser = new FormulaParser<ScalarValue, AstNode>(formula, lexer, new F());
        var f = parser.Formula();
        return (CellArea)f.Value;
    }
}