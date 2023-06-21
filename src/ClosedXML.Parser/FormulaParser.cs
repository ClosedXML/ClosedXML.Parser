using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Parser;

/// <summary>
/// A parser of Excel formulas, with main purpose of creating an abstract syntax tree.
/// </summary>
/// <remarks>
/// An implementation is a recursive descent parser, based on the ANTLR grammar.
/// </remarks>
/// <typeparam name="TScalarValue">Type of a scalar value used across expressions.</typeparam>
/// <typeparam name="TNode">Type of a node used in the AST.</typeparam>
public class FormulaParser<TScalarValue, TNode>
    where TNode : class
{
    private readonly string _input;
    private readonly FormulaLexer _tokenSource;
    private readonly IAstFactory<TScalarValue, TNode> _factory;

    // Current lookahead token index
    private int _la;

    public FormulaParser(string input, FormulaLexer tokenSource, IAstFactory<TScalarValue, TNode> factory)
    {
        _input = input;
        _tokenSource = tokenSource;
        _factory = factory;
        Consume();
    }

    /// <summary>
    /// Parse a formula.
    /// </summary>
    public TNode Formula()
    {
        if (_la == FormulaLexer.SPACE)
            Consume();
        var expression = Expression(false, out _);
        if (_la != FormulaLexer.Eof)
            throw new Exception("Expression is not completely parsed.");

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
                case FormulaLexer.GREATER_OR_EQUAL_THAN:
                    cmpOp = BinaryOperation.GreaterOrEqualThan;
                    break;
                case FormulaLexer.LESS_OR_EQUAL_THAN:
                    cmpOp = BinaryOperation.LessOrEqualThan;
                    break;
                case FormulaLexer.LESS_THAN:
                    cmpOp = BinaryOperation.LessThan;
                    break;
                case FormulaLexer.GREATER_THAN:
                    cmpOp = BinaryOperation.GreaterThan;
                    break;
                case FormulaLexer.NOT_EQUAL:
                    cmpOp = BinaryOperation.NotEqual;
                    break;
                case FormulaLexer.EQUAL:
                    cmpOp = BinaryOperation.Equal;
                    break;
                default:
                    return leftNode;
            }

            Consume();
            isPureRef = false;

            var rightNode = ConcatExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(cmpOp, leftNode, rightNode);
        }
    }

    private TNode ConcatExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var leftNode = AdditiveExpression(skipRangeUnion, out isPureRef);
        while (_la == FormulaLexer.CONCAT)
        {
            Consume();
            isPureRef = false;
            var rightNode = AdditiveExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(BinaryOperation.Concat, leftNode, rightNode);
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
                case FormulaLexer.PLUS:
                    op = BinaryOperation.Plus;
                    break;
                case FormulaLexer.MINUS:
                    op = BinaryOperation.Minus;
                    break;
                default:
                    return leftNode;
            }

            Consume();
            isPureRef = false;
            var rightNode = MultiplyingExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(op, leftNode, rightNode);
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
                case FormulaLexer.MULT:
                    op = BinaryOperation.Mult;
                    break;
                case FormulaLexer.DIV:
                    op = BinaryOperation.Div;
                    break;
                default:
                    return leftNode;
            }

            Consume();
            isPureRef = false;
            var appendNode = PowExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(op, leftNode, appendNode);
        }
    }

    private TNode PowExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var leftNode = PercentExpression(skipRangeUnion, out isPureRef);
        while (_la == FormulaLexer.POW)
        {
            Consume();
            isPureRef = false;
            var rightNode = PercentExpression(skipRangeUnion, out _);
            leftNode = _factory.BinaryNode(BinaryOperation.Pow, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode PercentExpression(bool skipRangeUnion, out bool isPureRef)
    {
        var prefixAtomNode = PrefixAtomExpression(skipRangeUnion, out isPureRef);
        if (_la == FormulaLexer.PERCENT)
        {
            Consume();
            isPureRef = false;
            return _factory.Unary('%', prefixAtomNode);
        }

        return prefixAtomNode;
    }

    private TNode PrefixAtomExpression(bool skipRangeUnion, out bool isPureRef)
    {
        char op;
        switch (_la)
        {
            case FormulaLexer.PLUS:
                op = '+';
                break;
            case FormulaLexer.MINUS:
                op = '+';
                break;
            default:
                return AtomExpression(skipRangeUnion, out isPureRef);
        }

        Consume();
        var neutralAtom = PrefixAtomExpression(skipRangeUnion, out _);
        isPureRef = false;
        return _factory.Unary(op, neutralAtom);
    }

    private TNode AtomExpression(bool skipRangeUnion, out bool isPureRef)
    {
        switch (_la)
        {
            // Constant
            case FormulaLexer.NONREF_ERRORS:
            case FormulaLexer.LOGICAL_CONSTANT:
            case FormulaLexer.NUMERICAL_CONSTANT:
            case FormulaLexer.STRING_CONSTANT:
            case FormulaLexer.OPEN_CURLY:
                isPureRef = false;
                var constantNode = Constant();
                return constantNode;

            // '(' expression ')'
            case FormulaLexer.OPEN_BRACE:
                Consume();
                var nestedNode = Expression(false, out isPureRef);
                Match(FormulaLexer.CLOSE_BRACE);

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
                        return RefIntersectionExpression(true, nestedNode);

                    return RefExpression(true, nestedNode);
                }

                return nestedNode;

            // function_call
            case FormulaLexer.CELL_FUNCTION_LIST:
            case FormulaLexer.USER_DEFINED_FUNCTION_NAME:
                isPureRef = false;
                var functionName = GetCurrentToken();
                Consume();
                var args = ArgumentList();
                return _factory.Function(functionName, args);

            // ref_expression
            default:
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
        var leftNode = RefIntersectionExpression(replaceFirstAtom, refAtom);
        while (_la == FormulaLexer.COMMA)
        {
            Consume();
            var rightNode = RefIntersectionExpression();
            leftNode = _factory.BinaryNode(BinaryOperation.Union, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode RefIntersectionExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var leftNode = RefRangeExpression(replaceFirstAtom, refAtom);
        while (_la == FormulaLexer.SPACE)
        {
            Consume();
            var rightNode = RefRangeExpression();
            leftNode = _factory.BinaryNode(BinaryOperation.Intersection, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode RefRangeExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var leftNode = RefAtomExpression(replaceFirstAtom, refAtom);
        while (_la == FormulaLexer.COLON)
        {
            Consume();
            var rightNode = RefAtomExpression();
            leftNode = _factory.BinaryNode(BinaryOperation.Range, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode RefAtomExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        // A backtracking of an incorrect detection whether an expression in a braces is value expression or ref expression.
        if (replaceFirstAtom)
            return refAtom!;

        switch (_la)
        {
            case FormulaLexer.REF_CONSTANT:
                var refError = _factory.ErrorNode(GetCurrentToken());
                Consume();
                return refError;

            case FormulaLexer.OPEN_BRACE:
                Consume();
                var refExpression = RefExpression();
                Match(FormulaLexer.CLOSE_BRACE);
                return refExpression;

            // cell_reference has been inlined into this switch

            // local_cell_reference
            case FormulaLexer.A1_REFERENCE:
                var referenceToken = GetCurrentToken();
                var cellArea = TokenParser.ParseA1Reference(referenceToken);
                var localCellReferenceNode = _factory.LocalCellReference(referenceToken, cellArea);
                Consume();
                return localCellReferenceNode;

            // external_cell_reference
            // case FormulaLexer.BANG_REFERENCE: Formula shouldn't contain BANG_REFERENCE, see grammar
            case FormulaLexer.SINGLE_SHEET_REFERENCE:
                var externalCellReferenceNode = _factory.ExternalCellReference(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                return externalCellReferenceNode;

            // external_cell_reference
            case FormulaLexer.SHEET_RANGE_PREFIX:
                var startCharIndex = _tokenSource.TokenStartCharIndex;
                Consume();
                Match(FormulaLexer.A1_REFERENCE);
                var sheetRangeReferenceNode = _factory.ExternalCellReference(_input, startCharIndex, _tokenSource.CharIndex - startCharIndex);
                return sheetRangeReferenceNode;

            // ref_function_call
            case FormulaLexer.REF_FUNCTION_LIST:
                var refFunctionName = GetCurrentToken();
                Consume();
                var args = ArgumentList();
                return _factory.Function(refFunctionName, args);

            // name_reference | structure_reference - all variants are expanded from the grammar.

            // Either defined name or table name for a structure reference
            case FormulaLexer.NAME:
                var localName = GetCurrentToken();
                Consume();
                if (_la == FormulaLexer.INTRA_TABLE_REFERENCE)
                {
                    var tableReference = GetCurrentToken();
                    Consume();
                    return _factory.StructureReference(localName, tableReference);
                }

                return _factory.LocalNameReference(localName);

            // reference to another workbook
            case FormulaLexer.BOOK_PREFIX:
                var bookPrefix = GetCurrentToken();
                Consume();
                var externalName = GetCurrentToken();
                Match(FormulaLexer.NAME);
                if (_la == FormulaLexer.INTRA_TABLE_REFERENCE)
                {
                    var intraTableReference = GetCurrentToken();
                    Consume();
                    return _factory.StructureReference(bookPrefix, externalName, intraTableReference);
                }

                return _factory.ExternalNameReference(bookPrefix, externalName);

            // name_reference - name in a specific sheet.
            case FormulaLexer.SINGLE_SHEET_PREFIX:
                var sheetPrefix = GetCurrentToken();
                Consume();
                var name = GetCurrentToken();
                Match(FormulaLexer.NAME);
                return _factory.LocalNameReference(sheetPrefix, name);

            // structure_reference - only for formulas directly in the table, e.g. totals row.
            case FormulaLexer.INTRA_TABLE_REFERENCE:
                var localTableReference = GetCurrentToken();
                Consume();
                return _factory.StructureReference(localTableReference);
        }

        throw UnexpectedTokenError();
    }

    private TNode Constant()
    {
        switch (_la)
        {
            case FormulaLexer.NONREF_ERRORS:
                var errorNode = _factory.ErrorNode(GetCurrentToken());
                Consume();
                return errorNode;

            case FormulaLexer.LOGICAL_CONSTANT:
                var logicalNode = ConvertLogical();
                Consume();
                return logicalNode;

            case FormulaLexer.NUMERICAL_CONSTANT:
                var numberNode = ConvertNumber();
                Consume();
                return numberNode;

            case FormulaLexer.STRING_CONSTANT:
                var textNode = ConvertText();
                Consume();
                return textNode;

            case FormulaLexer.OPEN_CURLY:
                Consume();
                var arrayElements = ConstantListRows(out var rows, out var columns);
                Match(FormulaLexer.CLOSE_CURLY);
                return _factory.ArrayNode(rows, columns, arrayElements);

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
        while (_la == FormulaLexer.SEMICOLON)
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
        while (_la == FormulaLexer.COMMA)
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
            case FormulaLexer.REF_CONSTANT:
            case FormulaLexer.NONREF_ERRORS:
                value = _factory.ErrorValue(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                break;
            case FormulaLexer.LOGICAL_CONSTANT:
                value = _factory.LogicalValue(GetTokenLogicalValue());
                break;
            case FormulaLexer.MINUS:
                Consume();
                if (_la != FormulaLexer.NUMERICAL_CONSTANT)
                    throw UnexpectedTokenError(FormulaLexer.NUMERICAL_CONSTANT);

                value = _factory.NumberValue(-ParseNumber(GetCurrentToken()));
                break;
            case FormulaLexer.NUMERICAL_CONSTANT:
                value = _factory.NumberValue(ParseNumber(GetCurrentToken()));
                break;
            case FormulaLexer.STRING_CONSTANT:
                value = ConvertTextValue();
                break;
            default:
                throw UnexpectedTokenError();
        }

        Consume();
        return value;
    }

    private List<TNode> ArgumentList()
    {
        var args = new List<TNode>();
        while (true)
        {
            if (_la == FormulaLexer.CLOSE_BRACE)
            {
                Consume();
                return args;
            }

            if (_la == FormulaLexer.COMMA)
            {
                Consume();
                args.Add(_factory.BlankNode());
            }
            else
            {
                var arg = Expression(true, out _);
                args.Add(arg);
                if (_la == FormulaLexer.CLOSE_BRACE)
                {
                    Consume();
                    return args;
                }

                Match(FormulaLexer.COMMA);
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
        var token = _tokenSource.NextToken();
        _la = token.Type;
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
        return _factory.LogicalNode(GetTokenLogicalValue());
    }

    private bool GetTokenLogicalValue()
    {
        return _input[_tokenSource.TokenStartCharIndex] is 'T' or 't';
    }

    private TNode ConvertNumber()
    {
        var number = ParseNumber(GetCurrentToken());
        var numberNode = _factory.NumberNode(number);
        return numberNode;
    }

    private TNode ConvertText()
    {
        var token = GetCurrentToken();
        Span<char> buffer = stackalloc char[token.Length];
        return ConvertTextValue(GetCurrentToken(), out var slice, ref buffer)
            ? _factory.TextNode(slice)
            : _factory.TextNode(buffer);
    }

    private TScalarValue ConvertTextValue()
    {
        var token = GetCurrentToken();
        Span<char> buffer = stackalloc char[token.Length];
        return ConvertTextValue(GetCurrentToken(), out var slice, ref buffer)
            ? _factory.TextValue(slice)
            : _factory.TextValue(buffer);
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
        return new Exception($"Error at line {_tokenSource.Line}:{_tokenSource.Column} of '{_input}': {message}");
    }

    private ReadOnlySpan<char> GetCurrentToken()
    {
        return _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
    }

    private static string GetTokenName(int tokenType) => FormulaLexer.ruleNames[tokenType - 1];

    private string GetLaTokenName() => GetTokenName(_la);
}
