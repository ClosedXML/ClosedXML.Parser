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

    private class ReferenceVisitor : BaseVisitor
    {
        public override string Reference(Result context, SymbolRange range, ReferenceArea reference)
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

        public virtual TNode SheetReference(TContext context, string sheet, ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode Reference3D(TContext context, string firstSheet, string lastSheet, ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalSheetReference(TContext context, int workbookIndex, string sheet, ReferenceArea reference)
        {
            return _defaultNode;
        }

        public virtual TNode ExternalReference3D(TContext context, int workbookIndex, string firstSheet, string lastSheet,
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

        public virtual TNode CellFunction(TContext context, RowCol cell, IReadOnlyList<TNode> arguments)
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