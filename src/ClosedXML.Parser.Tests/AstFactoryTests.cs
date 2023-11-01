namespace ClosedXML.Parser.Tests;

public class AstFactoryTests
{
    [Theory]
    [InlineData("A1", 0, 2)]
    [InlineData("A1 ", 0, 2)]
    [InlineData(" A1 ", 1, 3)]
    [InlineData(" $B7:D$18 ", 1, 9)]
    [InlineData("  1:7", 2, 5)]
    [InlineData("SUM(A:C)", 4, 7)]
    public void ReferenceRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ReferenceVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("Sheet!A1", 0, 8)]
    [InlineData(" S!$A$1:$B4 ", 1, 11)]
    [InlineData("1+'Johnny''s'!Z26", 2, 17)]
    public void SheetReferenceRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new SheetReferenceVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("Jan:Feb!A1", 0, 10)]
    [InlineData("1+Zara:Beta!$A$1:$B4+4", 2, 20)]
    [InlineData("1+'2022 Q1:2024 Q1'!Z26", 2, 23)]
    public void Reference3DRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new Reference3DVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("1+[2]S!A1 + 2", 2, 9)]
    [InlineData("1+'[2]D and D'!A1*2", 2, 17)]
    public void ExternalSheetReferenceRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ExternalSheetReferenceVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("1+[2]A:Z!A1 + 2", 2, 11)]
    [InlineData("1+'[2]D and D:B and B'!A1*2", 2, 25)]
    public void ExternalReference3DRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ExternalReference3DVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("2+A1(14,7)+2", 2, 10)]
    public void CellFunctionRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new CellFunctionVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("2+[Column]", 2, 10)]
    public void StructureReferenceNoTableRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new StructureReferenceNoTableVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("2+Table1[Column]+3", 2, 16)]
    public void StructureReferenceRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new StructureReferenceVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("2+[1]!Table1[Column]+3", 2, 20)]
    public void ExternalStructureReferenceRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ExternalStructureReferenceVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("A1+SUM(4)+name", 3, 9)]
    public void FunctionRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new FunctionVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("A1+[1]!SUM(4)+name", 3, 13)]
    public void ExternalFunctionRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ExternalFunctionVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("A1+Sheet!SUM(4)+name", 3, 15)]
    public void SheetFunctionRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new SheetFunctionVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("A1+[5]Sheet!SUM(4)+name", 3, 18)]
    public void ExternalSheetFunctionRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ExternalSheetFunctionVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("A1+TRUE-name", 8, 12)]
    public void NameRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new NameVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("Sheet!name + 4", 0, 10)]
    public void SheetNameRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new SheetNameVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("+[7]!name + 4", 1, 9)]
    public void ExternalNameRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ExternalNameVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Theory]
    [InlineData("+[7]Sheet!name + 4", 1, 14)]
    public void ExternalSheetNameRange(string formula, int start, int end)
    {
        var result = new Result();
        FormulaParser<object?, string, Result>.CellFormulaA1(formula, result, new ExternalSheetNameVisitor());
        Assert.Equal(new SymbolRange(start, end), result.Value);
    }

    [Fact]
    public void BinaryOperationRange()
    {
        var result = new List<SymbolRange>();
        FormulaParser<object?, string, List<SymbolRange>>.CellFormulaA1("1+2+3+4", result, new BinaryOperationVisitor());
        Assert.Equal(new[] { new SymbolRange(0, 3), new SymbolRange(0, 5), new SymbolRange(0, 7) }, result);
    }

    [Fact]
    public void UnaryOperationRange()
    {
        var result = new List<SymbolRange>();
        FormulaParser<object?, string, List<SymbolRange>>.CellFormulaA1("-7+8%", result, new UnaryOperationVisitor());
        Assert.Equal(new[] { new SymbolRange(0, 2), new SymbolRange(3, 5) }, result);
    }

    [Fact]
    public void NestedRange()
    {
        var result = new List<SymbolRange>();
        FormulaParser<object?, string, List<SymbolRange>>.CellFormulaA1("-( 1 + (2))+8%", result, new NestedOperationVisitor());
        Assert.Equal(new[] { new SymbolRange(7, 10), new SymbolRange(1, 11) }, result);
    }

    private class ReferenceVisitor : BaseVisitor
    {
        public override string Reference(Result context, SymbolRange range, ReferenceArea reference)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class SheetReferenceVisitor : BaseVisitor
    {
        public override string SheetReference(Result context, SymbolRange range, string sheet, ReferenceArea reference)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class Reference3DVisitor : BaseVisitor
    {
        public override string Reference3D(Result context, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalSheetReferenceVisitor : BaseVisitor
    {
        public override string ExternalSheetReference(Result context, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalReference3DVisitor : BaseVisitor
    {
        public override string ExternalReference3D(Result context, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class CellFunctionVisitor : BaseVisitor
    {
        public override string CellFunction(Result context, SymbolRange range, RowCol cell, IReadOnlyList<string> arguments)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class StructureReferenceNoTableVisitor : BaseVisitor
    {
        public override string StructureReference(Result context, SymbolRange range, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class StructureReferenceVisitor : BaseVisitor
    {
        public override string StructureReference(Result context, SymbolRange range, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalStructureReferenceVisitor : BaseVisitor
    {
        public override string ExternalStructureReference(Result context, SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class FunctionVisitor : BaseVisitor
    {
        public override string Function(Result context, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalFunctionVisitor : BaseVisitor
    {
        public override string ExternalFunction(Result context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class SheetFunctionVisitor : BaseVisitor
    {
        public override string Function(Result context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalSheetFunctionVisitor : BaseVisitor
    {
        public override string ExternalFunction(Result context, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<string> arguments)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class NameVisitor : BaseVisitor
    {
        public override string Name(Result context, SymbolRange range, string name)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class SheetNameVisitor : BaseVisitor
    {
        public override string SheetName(Result context, SymbolRange range, string sheetName, string name)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalNameVisitor : BaseVisitor
    {
        public override string ExternalName(Result context, SymbolRange range, int workbookIndex, string name)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class ExternalSheetNameVisitor : BaseVisitor
    {
        public override string ExternalSheetName(Result context, SymbolRange range, int workbookIndex, string sheet, string name)
        {
            context.Value = range;
            return string.Empty;
        }
    }

    private class BinaryOperationVisitor : BaseVisitor<object?, string, List<SymbolRange>>
    {
        public BinaryOperationVisitor()
            : base(null, string.Empty)
        {
        }

        public override string BinaryNode(List<SymbolRange> context, SymbolRange range, BinaryOperation operation, string leftNode, string rightNode)
        {
            context.Add(range);
            return string.Empty;
        }
    }

    private class UnaryOperationVisitor : BaseVisitor<object?, string, List<SymbolRange>>
    {
        public UnaryOperationVisitor()
            : base(null, string.Empty)
        {
        }

        public override string Unary(List<SymbolRange> context, SymbolRange range, UnaryOperation operation, string node)
        {
            context.Add(range);
            return string.Empty;
        }
    }

    private class NestedOperationVisitor : BaseVisitor<object?, string, List<SymbolRange>>
    {
        public NestedOperationVisitor()
            : base(null, string.Empty)
        {
        }

        public override string Nested(List<SymbolRange> context, SymbolRange range, string node)
        {
            context.Add(range);
            return string.Empty;
        }
    }

    private class BaseVisitor : BaseVisitor<object?, string, Result>
    {
        protected BaseVisitor()
            : base(null, string.Empty)
        {
        }
    }

    private class BaseVisitor<TScalarValue, TNode, TContext> : IAstFactory<TScalarValue, TNode, TContext>
        where TNode : class
    {
        private readonly TScalarValue _defaultScalar;
        private readonly TNode _defaultNode;

        protected BaseVisitor(TScalarValue defaultScalar, TNode defaultNode)
        {
            _defaultScalar = defaultScalar;
            _defaultNode = defaultNode;
        }

        public virtual TScalarValue LogicalValue(TContext context, bool value)
        {
            return _defaultScalar;
        }

        public virtual TScalarValue NumberValue(TContext context, double value)
        {
            return _defaultScalar;
        }

        public virtual TScalarValue TextValue(TContext context, string text)
        {
            return _defaultScalar;
        }

        public virtual TScalarValue ErrorValue(TContext context, ReadOnlySpan<char> error)
        {
            return _defaultScalar;
        }

        public virtual TNode ArrayNode(TContext context, int rows, int columns, IReadOnlyList<TScalarValue> elements)
        {
            return _defaultNode;
        }

        public virtual TNode BlankNode(TContext context)
        {
            return _defaultNode;
        }

        public virtual TNode LogicalNode(TContext context, bool value)
        {
            return _defaultNode;
        }

        public virtual TNode ErrorNode(TContext context, ReadOnlySpan<char> error)
        {
            return _defaultNode;
        }

        public virtual TNode NumberNode(TContext context, double value)
        {
            return _defaultNode;
        }

        public virtual TNode TextNode(TContext context, string text)
        {
            return _defaultNode;
        }

        public virtual TNode Reference(TContext context, SymbolRange range, ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode SheetReference(TContext context, SymbolRange range, string sheet, ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode Reference3D(TContext context, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalSheetReference(TContext context, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalReference3D(TContext context, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet,
            ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode Function(TContext context, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode Function(TContext context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> args)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalFunction(TContext context, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalFunction(TContext context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode CellFunction(TContext context, SymbolRange range, RowCol cell, IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode StructureReference(TContext context, SymbolRange range, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            return _defaultNode;
        }

        public virtual TNode StructureReference(TContext context, SymbolRange range, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalStructureReference(TContext context, SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn)
        {
            return _defaultNode;
        }

        public virtual TNode Name(TContext context, SymbolRange range, string name)
        {
            return _defaultNode;
        }

        public virtual TNode SheetName(TContext context, SymbolRange range, string sheet, string name)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalName(TContext context, SymbolRange range, int workbookIndex, string name)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalSheetName(TContext context, SymbolRange range, int workbookIndex, string sheet, string name)
        {
            return _defaultNode;
        }

        public virtual TNode BinaryNode(TContext context, SymbolRange range, BinaryOperation operation, TNode leftNode, TNode rightNode)
        {
            return _defaultNode;
        }

        public virtual TNode Unary(TContext context, SymbolRange range, UnaryOperation operation, TNode node)
        {
            return _defaultNode;
        }

        public virtual TNode Nested(TContext context, SymbolRange range, TNode node)
        {
            return _defaultNode;
        }
    }

    private class Result
    {
        internal SymbolRange? Value { get; set; }
    }
}