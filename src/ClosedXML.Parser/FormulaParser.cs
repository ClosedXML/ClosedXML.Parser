using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Parser.Rolex;

namespace ClosedXML.Parser;

/// <summary>
/// A parser of Excel formulas, with main purpose of creating an abstract syntax tree.
/// </summary>
/// <remarks>
/// An implementation is a recursive descent parser, based on the ANTLR grammar.
/// </remarks>
/// <typeparam name="TScalarValue">Type of a scalar value used across expressions.</typeparam>
/// <typeparam name="TNode">Type of a node used in the AST.</typeparam>
/// <typeparam name="TContext">A context of the parsing. It's passed to every factory method and can contain global info that doesn't belong individual nodes.</typeparam>
public class FormulaParser<TScalarValue, TNode, TContext>
    where TNode : class
{
    private readonly string _input;
    private readonly List<Token> _tokens;
    private readonly IAstFactory<TScalarValue, TNode, TContext> _factory;
    private readonly TContext _context;

    /// <summary>
    /// Is parser in A1 mode (true) or R1C1 mode (false)?
    /// </summary>
    private readonly bool _a1Mode;
    private Token _tokenSource;
    private int _tokenIndex = -1;

    // Current lookahead token index
    private int _la;

    private FormulaParser(string formula, TContext context, IAstFactory<TScalarValue, TNode, TContext> factory, bool a1Mode)
    {
        // Trim the end, so ref_intersection_expression that tried to parse SPACE as an operator
        // doesn't recognize spaces at the end of formula as operators. The control tokens of
        // the formula have whitespaces around them (unlike params), so the whitespaces should
        // be consumed by control tokens (e.g. ` IF ( A1 ) ` will be split into `IF ( `, `A1` and ` ) `)
        // but to avoid the whitespace at the end, trim it.
        var trimmedFormula = formula.AsSpan().TrimEnd();
        _input = formula;
        _context = context;
        _tokens = a1Mode
            ? RolexLexer.GetTokensA1(trimmedFormula)
            : RolexLexer.GetTokensR1C1(trimmedFormula);
        _factory = factory;
        _a1Mode = a1Mode;
        Consume();
    }

    /// <summary>
    /// Parse a formula using A1 semantic for references. 
    /// </summary>
    /// <param name="formula">Formula text that will be parsed.</param>
    /// <param name="context">Context that is going to be passed to every method of the <paramref name="factory"/>.</param>
    /// <param name="factory">Factory to create nodes of AST tree.</param>
    /// <exception cref="ParsingException">If the formula doesn't satisfy the grammar.</exception>
    public static TNode CellFormulaA1(string formula, TContext context, IAstFactory<TScalarValue, TNode, TContext> factory)
    {
        var parser = new FormulaParser<TScalarValue, TNode, TContext>(formula, context, factory, true);
        return parser.Formula();
    }

    /// <summary>
    /// Parse a formula using R1C1 semantic for references. 
    /// </summary>
    /// <param name="formula">Formula text that will be parsed.</param>
    /// <param name="context">Context that is going to be passed to every method of the <paramref name="factory"/>.</param>
    /// <param name="factory">Factory to create nodes of AST tree.</param>
    /// <exception cref="ParsingException">If the formula doesn't satisfy the grammar.</exception>
    public static TNode CellFormulaR1C1(string formula, TContext context, IAstFactory<TScalarValue, TNode, TContext> factory)
    {
        var parser = new FormulaParser<TScalarValue, TNode, TContext>(formula, context, factory, false);
        return parser.Formula();
    }

    private TNode Formula()
    {
        if (_tokens[_tokens.Count - 1].SymbolId == Token.ErrorSymbolId)
            throw new ParsingException($"Unable to determine token for '{_input}' at index {_tokens[_tokens.Count - 1].StartIndex}.");

        if (_la == Token.SPACE)
            Consume();
        var expression = Expression(false, out _);
        if (_la != Token.EofSymbolId)
        {
            var parsedPart = _input.Substring(0, _tokenSource.StartIndex);
            var remainder = _input.Substring(_tokenSource.StartIndex);
            throw new ParsingException($"The formula `{_input}` wasn't parsed correctly. The expression `{parsedPart}` was parsed, but the rest `{remainder}` wasn't.");
        }

        return expression;
    }

    private TNode Expression(bool skipRangeUnion, out bool isPureRef)
    {
        var leftNode = ConcatExpression(skipRangeUnion, out isPureRef);
        while (true)
        {
            BinaryOperation cmpOp;
            switch (_la)
            {
                case Token.GREATER_OR_EQUAL_THAN:
                    cmpOp = BinaryOperation.GreaterOrEqualThan;
                    break;
                case Token.LESS_OR_EQUAL_THAN:
                    cmpOp = BinaryOperation.LessOrEqualThan;
                    break;
                case Token.LESS_THAN:
                    cmpOp = BinaryOperation.LessThan;
                    break;
                case Token.GREATER_THAN:
                    cmpOp = BinaryOperation.GreaterThan;
                    break;
                case Token.NOT_EQUAL:
                    cmpOp = BinaryOperation.NotEqual;
                    break;
                case Token.EQUAL:
                    cmpOp = BinaryOperation.Equal;
                    break;
                default:
                    return leftNode;
            }

            Consume();
            isPureRef = false;

            var rightNode = ConcatExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(_context, cmpOp, leftNode, rightNode);
        }
    }

    private TNode ConcatExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var leftNode = AdditiveExpression(skipRangeUnion, out isPureRef);
        while (_la == Token.CONCAT)
        {
            Consume();
            isPureRef = false;
            var rightNode = AdditiveExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(_context, BinaryOperation.Concat, leftNode, rightNode);
        }

        return leftNode;
    }
    private TNode AdditiveExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var leftNode = MultiplyingExpression(skipRangeUnion, out isPureRef);
        while (true)
        {
            BinaryOperation op;
            switch (_la)
            {
                case Token.PLUS:
                    op = BinaryOperation.Addition;
                    break;
                case Token.MINUS:
                    op = BinaryOperation.Subtraction;
                    break;
                default:
                    return leftNode;
            }

            Consume();
            isPureRef = false;
            var rightNode = MultiplyingExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(_context, op, leftNode, rightNode);
        }
    }

    private TNode MultiplyingExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var leftNode = PowExpression(skipRangeUnion, out isPureRef);
        while (true)
        {
            BinaryOperation op;
            switch (_la)
            {
                case Token.MULT:
                    op = BinaryOperation.Multiplication;
                    break;
                case Token.DIV:
                    op = BinaryOperation.Division;
                    break;
                default:
                    return leftNode;
            }

            Consume();
            isPureRef = false;
            var appendNode = PowExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(_context, op, leftNode, appendNode);
        }
    }

    private TNode PowExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var leftNode = PercentExpression(skipRangeUnion, out isPureRef);
        while (_la == Token.POW)
        {
            Consume();
            isPureRef = false;
            var rightNode = PercentExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(_context, BinaryOperation.Power, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode PercentExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var prefixAtomNode = PrefixAtomExpression(skipRangeUnion, out isPureRef);
        var percentNode = prefixAtomNode;
        while (_la == Token.PERCENT)
        {
            Consume();
            isPureRef = false;
            percentNode = _factory.Unary(_context, UnaryOperation.Percent, percentNode);
        }

        return percentNode;
    }

    /// <summary>
    /// Parser for two rules unified into a single method.
    /// <para>
    /// <c>
    /// prefix_atom_expression
    ///     : (PLUS | MINUS) prefix_atom_expression
    ///     | atom_expression
    ///     ;
    /// </c>
    /// 
    /// <c>
    /// arg_prefix_atom_expression
    ///     : (PLUS | MINUS) arg_prefix_atom_expression
    ///     | arg_atom_expression
    ///     ;     
    /// </c>
    /// </para>
    /// </summary>
    /// <param name="skipRangeUnion">Does the method represent <c>prefix_atom_expression</c> (<c>false</c>) or <c>arg_prefix_atom_expression</c> (<c>true</c>)</param>
    /// <param name="isPureRef">Is the expression of the node a reference expression?</param>
    /// <returns></returns>
    private TNode PrefixAtomExpression(bool skipRangeUnion, out bool isPureRef)
    {
        UnaryOperation op;
        switch (_la)
        {
            case Token.PLUS:
                op = UnaryOperation.Plus;
                break;
            case Token.MINUS:
                op = UnaryOperation.Minus;
                break;
            default:
                return AtomExpression(skipRangeUnion, out isPureRef);
        }

        Consume();
        var neutralAtom = PrefixAtomExpression(skipRangeUnion, out _);
        isPureRef = false;
        return _factory.Unary(_context, op, neutralAtom);
    }

    private TNode AtomExpression(bool skipRangeUnion, out bool isPureRef)
    {
        switch (_la)
        {
            // Constant
            case Token.NONREF_ERRORS:
            case Token.LOGICAL_CONSTANT:
            case Token.NUMERICAL_CONSTANT:
            case Token.STRING_CONSTANT:
            case Token.OPEN_CURLY:
                isPureRef = false;
                var constantNode = Constant();
                return constantNode;

            // '(' expression ')'
            case Token.OPEN_BRACE:
                Consume();
                var nestedNode = _factory.Nested(_context, Expression(false, out isPureRef));
                Match(Token.CLOSE_BRACE);

                // This is the point of an ambiguity. Atom should be a value, but it can
                // be determined by calling an expression inside the braces or
                // it can go through ref_expression path that also has an ref_expression
                // in braces. The second option should be seriously rare, so parser
                // tracks whether expression is a ref_expression and if it is, we 'backtrack'
                // through patching to the correct path of the 'ref_expression' below.
                // Example: '(A1) A1:B2' <- the '(A1)' should be detected as ref_expression
                // and thus the ' ' intersection operator be valid. Of course, braces can be
                // very nested '(((A1))) A1:B2' and when we are entering the brace, there is
                // no way to detect whether it is ref_expression or expression.
                if (isPureRef)
                {
                    // Incorrect expectation, backtrack to the ref_expression
                    // note the passed true argument for 'replaceFirstAtom'
                    if (skipRangeUnion)
                        return RefImplicitExpression(true, nestedNode);

                    return RefExpression(true, nestedNode);
                }

                return nestedNode;

            // function_call
            case Token.CELL_FUNCTION_LIST:
                {
                    isPureRef = false;
                    var cellReference = TokenParser.ExtractCellFunction(GetCurrentToken());
                    Consume();
                    var args = ArgumentList();
                    return _factory.CellFunction(_context, cellReference, args);
                }

            case Token.USER_DEFINED_FUNCTION_NAME:
                isPureRef = false;
                return LocalFunctionCall();

            default:
                // function_call : SINGLE_SHEET_PREFIX USER_DEFINED_FUNCTION_NAME argument_list
                if (_la == Token.SINGLE_SHEET_PREFIX && LL(1) == Token.USER_DEFINED_FUNCTION_NAME)
                {
                    isPureRef = false;
                    TokenParser.ParseSingleSheetPrefix(GetCurrentToken(), out var wbIndex, out var sheetName);
                    Consume();
                    var functionName = TokenParser.ExtractLocalFunctionName(GetCurrentToken());
                    Consume();
                    var args = ArgumentList();
                    return wbIndex is null
                        ? _factory.Function(_context, sheetName, functionName, args)
                        : _factory.ExternalFunction(_context, wbIndex.Value, sheetName, functionName, args);
                }

                // function_call : BOOK_PREFIX USER_DEFINED_FUNCTION_NAME argument_list
                if (_la == Token.BOOK_PREFIX && LL(1) == Token.USER_DEFINED_FUNCTION_NAME)
                {
                    isPureRef = false;
                    var wbIndex = TokenParser.ParseBookPrefix(GetCurrentToken());
                    Consume();
                    var functionName = TokenParser.ExtractLocalFunctionName(GetCurrentToken());
                    Consume();
                    var args = ArgumentList();
                    return _factory.ExternalFunction(_context, wbIndex, functionName, args);
                }

                // ref_expression 
                if (skipRangeUnion)
                {
                    isPureRef = true;
                    return RefIntersectionExpression();
                }

                isPureRef = true;
                return RefExpression();
        }
    }

    private TNode RefExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var leftNode = RefImplicitExpression(replaceFirstAtom, refAtom);
        while (_la == Token.COMMA)
        {
            Consume();
            var rightNode = RefImplicitExpression();
            leftNode = _factory.BinaryNode(_context, BinaryOperation.Union, leftNode, rightNode);
        }

        return leftNode;
    }

    /// <summary>
    /// <code>
    /// ref_implicit_expression
    ///        : INTERSECT ref_implicit_expression
    ///        | ref_intersection_expression
    ///        ;
    /// </code>
    /// </summary>
    private TNode RefImplicitExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        if (_la == Token.INTERSECT)
        {
            Consume();
            var refNode = RefImplicitExpression(replaceFirstAtom, refAtom);
            return _factory.Unary(_context, UnaryOperation.ImplicitIntersection, refNode);
        }

        return RefIntersectionExpression(replaceFirstAtom, refAtom);
    }

    private TNode RefIntersectionExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var leftNode = RefRangeExpression(replaceFirstAtom, refAtom);
        while (_la == Token.SPACE)
        {
            Consume();
            var rightNode = RefRangeExpression();
            leftNode = _factory.BinaryNode(_context, BinaryOperation.Intersection, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode RefRangeExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var leftNode = RefSpillExpression(replaceFirstAtom, refAtom);
        while (_la == Token.COLON)
        {
            Consume();
            var rightNode = RefSpillExpression();
            leftNode = _factory.BinaryNode(_context, BinaryOperation.Range, leftNode, rightNode);
        }

        return leftNode;
    }

    /// <summary>
    /// Parser of the following node.
    /// <c>
    /// ref_spill_expression
    ///     : ref_atom_expression SPILL?
    ///     ;
    /// </c>
    /// </summary>
    private TNode RefSpillExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var refAtomNode = RefAtomExpression(replaceFirstAtom, refAtom);
        if (_la == Token.SPILL)
        {
            Consume();
            return _factory.Unary(_context, UnaryOperation.SpillRange, refAtomNode);
        }

        return refAtomNode;
    }

    private TNode RefAtomExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        // A backtracking of an incorrect detection whether an expression in a braces is value expression or ref expression.
        if (replaceFirstAtom)
            return refAtom!;

        switch (_la)
        {
            case Token.REF_CONSTANT:
                return ErrorNode();

            case Token.OPEN_BRACE:
                Consume();
                var refExpression = _factory.Nested(_context, RefExpression());
                Match(Token.CLOSE_BRACE);
                return refExpression;

            // cell_reference has been inlined into this switch

            // cell_reference:
            //     (A1_CELL | A1_CELL COLON A1_CELL) -- inlined a1_reference
            case Token.A1_CELL:
                {
                    var startIdx = _tokenSource.StartIndex;
                    var area = TokenParser.ParseReference(GetCurrentToken(), _a1Mode);
                    Consume();
                    if (_la == Token.COLON && LL(1) == Token.A1_CELL)
                    {
                        Consume();
                        var secondCell = TokenParser.ParseReference(GetCurrentToken(), _a1Mode);
                        Consume();
                        area = new ReferenceArea(area.First, secondCell.First);
                    }

                    var endIdx = _tokenSource.StartIndex;
                    var reference = _factory.Reference(_context, new SymbolRange(startIdx, endIdx), area);
                    return reference;
                }

            // cell_reference:
            //     (A1_SPAN_REFERENCE)  -- inlined a1_reference
            case Token.A1_SPAN_REFERENCE:
                {
                    var start = _tokenSource.StartIndex;
                    var area = TokenParser.ParseReference(GetCurrentToken(), _a1Mode);
                    Consume();
                    var end = _tokenSource.StartIndex;
                    var reference = _factory.Reference(_context, new SymbolRange(start, end), area);
                    return reference;
                }

            // external_cell_reference
            // case FormulaLexer.BANG_REFERENCE: Formula shouldn't contain BANG_REFERENCE, see grammar
            // external_cell_reference: SHEET_RANGE_PREFIX (A1_CELL | A1_CELL COLON A1_CELL | A1_SPAN_REFERENCE)
            case Token.SHEET_RANGE_PREFIX:
                {
                    var sheetRangePrefixToken = GetCurrentToken();
                    TokenParser.ParseSheetRangePrefix(sheetRangePrefixToken, out var wbIdx, out var firstName,
                        out var secondName);
                    Consume();

                    var area = A1Reference();
                    if (area is null)
                        throw UnexpectedTokenError(Token.A1_CELL, Token.A1_SPAN_REFERENCE);

                    return wbIdx is not null
                        ? _factory.ExternalReference3D(_context, wbIdx.Value, firstName, secondName, area.Value)
                        : _factory.Reference3D(_context, firstName, secondName, area.Value);
                }

            // ref_function_call
            case Token.REF_FUNCTION_LIST:
                return LocalFunctionCall();

            // name_reference | structure_reference - all variants are expanded from the grammar.

            // Either defined name or table name for a structure reference
            case Token.NAME:
                {
                    var localName = GetCurrentToken();
                    Consume();
                    if (_la == Token.INTRA_TABLE_REFERENCE)
                    {
                        TokenParser.ParseIntraTableReference(GetCurrentToken(), out var specifics, out var firstColumn, out var lastColumn);
                        Consume();
                        return _factory.StructureReference(_context, localName.ToString(), specifics, firstColumn, lastColumn ?? firstColumn);
                    }

                    // 3D reference
                    if (_la == Token.COLON && LL(1) == Token.SINGLE_SHEET_PREFIX)
                    {
                        var firstSheetName = localName.ToString();
                        Consume(); // COLON

                        // TODO: Decouple book prefix from single sheet prefix
                        TokenParser.ParseSingleSheetPrefix(GetCurrentToken(), out var wbIdx, out string lastSheetName);
                        if (wbIdx is not null)
                            throw Error("External workbook not expected.");

                        Consume(); // SINGLE_SHEET_PREFIX

                        // After prefix, there must be A1Reference
                        var area = A1Reference();
                        if (area is null)
                            throw UnexpectedTokenError(Token.A1_CELL, Token.A1_SPAN_REFERENCE);

                        return _factory.Reference3D(_context, firstSheetName, lastSheetName, area.Value);
                    }

                    return _factory.Name(_context, localName.ToString());
                }

            // reference to another workbook
            case Token.BOOK_PREFIX:
                {
                    var bookPrefix = TokenParser.ParseBookPrefix(GetCurrentToken());
                    Consume();
                    var externalName = GetCurrentToken();
                    Match(Token.NAME);
                    if (_la == Token.INTRA_TABLE_REFERENCE)
                    {
                        TokenParser.ParseIntraTableReference(GetCurrentToken(), out var specifics, out var firstColumn, out var lastColumn);
                        Consume();
                        return _factory.ExternalStructureReference(_context, bookPrefix, externalName.ToString(), specifics, firstColumn, lastColumn ?? firstColumn);
                    }

                    return _factory.ExternalName(_context, bookPrefix, externalName.ToString());
                }
            // name_reference: SINGLE_SHEET_PREFIX NAME
            // external_cell_reference: SINGLE_SHEET_PREFIX (A1_CELL | A1_CELL COLON A1_CELL | A1_SPAN_REFERENCE | REF_CONSTANT)
            case Token.SINGLE_SHEET_PREFIX:
                {
                    var sheetPrefix = GetCurrentToken();
                    TokenParser.ParseSingleSheetPrefix(sheetPrefix, out var wbIdx, out string sheetName);
                    Consume();

                    var area = A1Reference();
                    if (area is not null)
                    {
                        return wbIdx is null
                            ? _factory.SheetReference(_context, sheetName, area.Value)
                            : _factory.ExternalSheetReference(_context, wbIdx.Value, sheetName, area.Value);
                    }

                    if (_la == Token.REF_CONSTANT)
                    {
                        var errorReference = _factory.ErrorNode(_context, GetCurrentToken()); // Sheet1!#REF! is a valid
                        Consume();
                        return errorReference;
                    }

                    // name_reference
                    var name = GetCurrentToken();
                    Match(Token.NAME);
                    return wbIdx is null
                        ? _factory.SheetName(_context, sheetName, name.ToString())
                        : _factory.ExternalSheetName(_context, wbIdx.Value, sheetName, name.ToString());
                }

            // structure_reference - only for formulas directly in the table, e.g. totals row.
            case Token.INTRA_TABLE_REFERENCE:
                {
                    var localTableReference = GetCurrentToken();
                    TokenParser.ParseIntraTableReference(localTableReference, out var specifics, out var firstColumn, out var lastColumn);
                    Consume();
                    return _factory.StructureReference(_context, specifics, firstColumn, lastColumn ?? firstColumn);
                }
        }

        throw UnexpectedTokenError();
    }

    /// <summary>
    /// <code>
    /// a1_reference
    ///     : A1_CELL
    ///     | A1_CELL COLON A1_CELL
    ///     | A1_SPAN_REFERENCE
    ///     ;
    /// </code>
    /// </summary>
    private ReferenceArea? A1Reference()
    {
        if (_la == Token.A1_CELL)
        {
            var cellToken = GetCurrentToken();
            var cell = TokenParser.ParseReference(cellToken, _a1Mode);
            Consume();
            var area = cell;
            if (_la == Token.COLON && LL(1) == Token.A1_CELL)
            {
                Consume();
                var secondCell = TokenParser.ParseReference(GetCurrentToken(), _a1Mode);
                area = new ReferenceArea(cell.First, secondCell.First);
                Consume();
            }

            return area;
        }

        if (_la == Token.A1_SPAN_REFERENCE)
        {
            var area = TokenParser.ParseReference(GetCurrentToken(), _a1Mode);
            Consume();
            return area;
        }

        return null;
    }

    private TNode ErrorNode()
    {
        var errorToken = GetCurrentToken();
        Span<char> normalizedError = stackalloc char[errorToken.Length];
        errorToken.ToUpperInvariant(normalizedError);
        var refError = _factory.ErrorNode(_context, normalizedError);
        Consume();
        return refError;
    }

    private TNode Constant()
    {
        switch (_la)
        {
            case Token.NONREF_ERRORS:
                return ErrorNode();

            case Token.LOGICAL_CONSTANT:
                var logicalNode = ConvertLogical();
                Consume();
                return logicalNode;

            case Token.NUMERICAL_CONSTANT:
                var numberNode = ConvertNumber();
                Consume();
                return numberNode;

            case Token.STRING_CONSTANT:
                var textNode = ConvertText();
                Consume();
                return textNode;

            case Token.OPEN_CURLY:
                Consume();
                var arrayElements = ConstantListRows(out var rows, out var columns);
                Match(Token.CLOSE_CURLY);
                return _factory.ArrayNode(_context, rows, columns, arrayElements);

            default:
                throw UnexpectedTokenError();
        }
    }

    private List<TScalarValue> ConstantListRows(out int rows, out int columns)
    {
        // First use list with doubling strategy
        var elements = new List<TScalarValue>();
        var rowSize = ConstantListRow(elements);
        var height = 1;
        while (_la == Token.SEMICOLON)
        {
            Consume();
            var nextRowSize = ConstantListRow(elements);
            if (nextRowSize != rowSize)
                throw Error("Rows of an array don't have same size.");

            height++;
        }

        rows = height;
        columns = rowSize;
        return elements;
    }

    private int ConstantListRow(List<TScalarValue> arrayElements)
    {
        var origSize = arrayElements.Count;
        var arrayElement = ArrayConstant();
        arrayElements.Add(arrayElement);
        while (_la == Token.COMMA)
        {
            Consume();
            var nextElement = ArrayConstant();
            arrayElements.Add(nextElement);
        }

        return arrayElements.Count - origSize;
    }

    private TScalarValue ArrayConstant()
    {
        TScalarValue value;
        switch (_la)
        {
            case Token.REF_CONSTANT:
            case Token.NONREF_ERRORS:
                // Convert to upper case on stack, because length of an error is limited to ~20
                var errorToken = GetCurrentToken();
                Span<char> normalizedError = stackalloc char[errorToken.Length];
                errorToken.ToUpperInvariant(normalizedError);
                value = _factory.ErrorValue(_context, normalizedError);
                break;
            case Token.LOGICAL_CONSTANT:
                value = _factory.LogicalValue(_context, GetTokenLogicalValue());
                break;
            case Token.MINUS:
                Consume();
                if (_la != Token.NUMERICAL_CONSTANT)
                    throw UnexpectedTokenError(Token.NUMERICAL_CONSTANT);

                value = _factory.NumberValue(_context, -ParseNumber(GetCurrentToken()));
                break;
            case Token.PLUS:
                Consume();
                if (_la != Token.NUMERICAL_CONSTANT)
                    throw UnexpectedTokenError(Token.NUMERICAL_CONSTANT);

                value = _factory.NumberValue(_context, ParseNumber(GetCurrentToken()));
                break;
            case Token.NUMERICAL_CONSTANT:
                value = _factory.NumberValue(_context, ParseNumber(GetCurrentToken()));
                break;
            case Token.STRING_CONSTANT:
                value = ConvertTextValue();
                break;
            default:
                throw UnexpectedTokenError();
        }

        Consume();
        return value;
    }

    private IReadOnlyList<TNode> ArgumentList()
    {
        // A special case, there are no arguments
        if (_la == Token.CLOSE_BRACE)
        {
            Consume();
            return Array.Empty<TNode>();
        }

        var args = new List<TNode>();
        while (true)
        {
            // At the start of the loop, previous argument
            // should have been consumed with a comma.
            if (_la == Token.COMMA)
            {
                // If there is a comma, it means there are
                // two commas in a row and thus a blank argument.
                args.Add(_factory.BlankNode(_context));
                Consume();
            }
            else if (_la == Token.CLOSE_BRACE)
            {
                // if there is a brace, it means the previous
                // comma is immediately followed by a brace `,)`
                // thus there is a blank node and end of args.
                args.Add(_factory.BlankNode(_context));
                Consume();
                return args;
            }
            else
            {
                // Path for a non-blank argument.
                var arg = Expression(true, out _);
                args.Add(arg);
                if (_la == Token.CLOSE_BRACE)
                {
                    Consume();
                    return args;
                }

                // Each argument must be followed by a comma. 
                Match(Token.COMMA);
            }
        }
    }

    private void Match(int expected)
    {
        if (_la != expected)
            throw UnexpectedTokenError(expected);

        Consume();
    }

    private void Consume()
    {
        _tokenSource = _tokens[++_tokenIndex];
        _la = _tokenSource.SymbolId;
    }

    private int LL(int lookAhead)
    {
        var idx = _tokenIndex + lookAhead;
        return idx < _tokens.Count ? _tokens[idx].SymbolId : Token.EofSymbolId;
    }

    private static double ParseNumber(ReadOnlySpan<char> number)
    {
        return double.Parse(
#if NETSTANDARD2_1
            number,
#else
            number.ToString(),
#endif
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent,
            CultureInfo.InvariantCulture);
    }

    private TNode ConvertLogical()
    {
        return _factory.LogicalNode(_context, GetTokenLogicalValue());
    }

    private bool GetTokenLogicalValue()
    {
        return _input[_tokenSource.StartIndex] is 'T' or 't';
    }

    private TNode ConvertNumber()
    {
        var number = ParseNumber(GetCurrentToken());
        var numberNode = _factory.NumberNode(_context, number);
        return numberNode;
    }

    private TNode ConvertText()
    {
        var token = GetCurrentToken();
        Span<char> buffer = stackalloc char[token.Length];
        return ConvertTextValue(GetCurrentToken(), out var slice, ref buffer)
            ? _factory.TextNode(_context, slice.ToString())
            : _factory.TextNode(_context, buffer.ToString());
    }

    private TScalarValue ConvertTextValue()
    {
        var token = GetCurrentToken();
        Span<char> buffer = stackalloc char[token.Length];
        return ConvertTextValue(GetCurrentToken(), out var slice, ref buffer)
            ? _factory.TextValue(_context, slice.ToString())
            : _factory.TextValue(_context, buffer.ToString());
    }

    private static bool ConvertTextValue(ReadOnlySpan<char> token, out ReadOnlySpan<char> copy, ref Span<char> buffer)
    {
        var text = token.Slice(1, token.Length - 2);
        var indexOfDQuote = text.IndexOf('"');
        var textMustBeUnescaped = indexOfDQuote >= 0;
        if (!textMustBeUnescaped)
        {
            copy = text;
            return true;
        }

        Span<char> unescaped = buffer;
        var tail = unescaped;
        var quoteCount = 0;
        do
        {
            var quoteText = text.Slice(0, indexOfDQuote + 1);
            quoteText.CopyTo(tail);
            tail = tail.Slice(indexOfDQuote + 1);
            text = text.Slice(indexOfDQuote + 2);
            indexOfDQuote = text.IndexOf('"');
            quoteCount++;
        } while (indexOfDQuote >= 0);

        text.CopyTo(tail);
        buffer = unescaped.Slice(0, token.Length - 2 - quoteCount);
        copy = default;
        return false;
    }

    private TNode LocalFunctionCall()
    {
        var functionName = TokenParser.ExtractLocalFunctionName(GetCurrentToken());
        Consume();
        var args = ArgumentList();
        return _factory.Function(_context, functionName, args);
    }

    private Exception UnexpectedTokenError(params int[] expectedToken)
    {
        return Error($"Unexpected token {GetLaTokenName()}, expected one of {string.Join(",", expectedToken.Select(GetTokenName))}.");
    }

    private Exception UnexpectedTokenError()
    {
        return Error($"Unexpected token {GetLaTokenName()}.");
    }

    private Exception Error(string message)
    {
        return new ParsingException($"Error at char {_tokenSource.StartIndex} of '{_input}': {message}");
    }

    private ReadOnlySpan<char> GetCurrentToken()
    {
        return _input.AsSpan(_tokenSource.StartIndex, _tokenSource.Length);
    }

    private static string GetTokenName(int tokenType) => Token.GetSymbolName(tokenType);

    private string GetLaTokenName() => GetTokenName(_la);
}