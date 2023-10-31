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
    [InlineData("1+Z:B!$A$1:$B4+4", 2, 14)]
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

        public virtual TNode Function(TContext context, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode Function(TContext context, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> args)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalFunction(TContext context, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalFunction(TContext context, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode CellFunction(TContext context, SymbolRange range, RowCol cell, IReadOnlyList<TNode> arguments)
        {
            return _defaultNode;
        }

        public virtual TNode StructureReference(TContext context, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            return _defaultNode;
        }

        public virtual TNode StructureReference(TContext context, string table, StructuredReferenceArea area, string? firstColumn,
            string? lastColumn)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalStructureReference(TContext context, int workbookIndex, string table, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn)
        {
            return _defaultNode;
        }

        public virtual TNode Name(TContext context, string name)
        {
            return _defaultNode;
        }

        public virtual TNode SheetName(TContext context, string sheet, string name)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalName(TContext context, int workbookIndex, string name)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalSheetName(TContext context, int workbookIndex, string sheet, string name)
        {
            return _defaultNode;
        }

        public virtual TNode BinaryNode(TContext context, BinaryOperation operation, TNode leftNode, TNode rightNode)
        {
            return _defaultNode;
        }

        public virtual TNode Unary(TContext context, UnaryOperation operation, TNode node)
        {
            return _defaultNode;
        }

        public virtual TNode Nested(TContext context, TNode node)
        {
            return _defaultNode;
        }
    }

    private class Result
    {
        internal SymbolRange? Value { get; set; }
    }
}