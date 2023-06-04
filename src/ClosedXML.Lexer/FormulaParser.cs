using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Lexer;

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
        var expression = Expression(false).Node;
        if (_la != FormulaLexer.Eof)
            throw new Exception("Expression is not completely parsed.");

        return expression;
    }

    private (bool IsPureRef, TNode Node) Expression(bool skipRangeUnion)
    {
        var (isPureRef, currentNode) = ConcatExpression(skipRangeUnion);
        while (true)
        {
            char op;
            var la = _la;
            switch (la)
            {
                case FormulaLexer.GREATER_OR_EQUAL_THAN:
                    op = '≥';
                    break;
                case FormulaLexer.LESS_OR_EQUAL_THAN:
                    op = '≤';
                    break;
                case FormulaLexer.LESS_THAN:
                    op = '<';
                    break;
                case FormulaLexer.GREATER_THAN:
                    op = '>';
                    break;
                case FormulaLexer.NOT_EQUAL:
                    op = '≠';
                    break;
                case FormulaLexer.EQUAL:
                    op = '=';
                    break;
                default:
                    return (isPureRef, currentNode);
            }

            Consume();
            isPureRef = false;

            var appendNode = ConcatExpression(skipRangeUnion).Node;
            currentNode = _factory.BinaryNode(op, currentNode, appendNode);
        }
    }

    private (bool IsPureRef, TNode Node) ConcatExpression(bool skipRangeUnion)
    {
        var (isPureRef, currentNode) = AdditiveExpression(skipRangeUnion);
        for (var la = _la; la == FormulaLexer.CONCAT; la = _la)
        {
            Consume();
            isPureRef = false;
            var appendNode = AdditiveExpression(skipRangeUnion).Node;
            currentNode = _factory.BinaryNode('&', currentNode, appendNode);
        }

        return (isPureRef, currentNode);
    }
    private (bool IsPureRef, TNode Node) AdditiveExpression(bool skipRangeUnion)
    {
        var (isPureRef, currentNode) = MultiplyingExpression(skipRangeUnion);
        for (var la = _la; la is FormulaLexer.PLUS or FormulaLexer.MINUS; la = _la)
        {
            Consume();
            isPureRef = false;
            var appendNode = MultiplyingExpression(skipRangeUnion).Item2;
            currentNode = _factory.BinaryNode(la == FormulaLexer.PLUS ? '+' : '-', currentNode, appendNode);
        }

        return (isPureRef, currentNode);
    }

    private (bool IsPureRef, TNode Node) MultiplyingExpression(bool skipRangeUnion)
    {
        var (isPureRef, currentNode) = PowExpression(skipRangeUnion);
        for (var la = _la; la is FormulaLexer.MULT or FormulaLexer.DIV; la = _la)
        {
            Consume();
            isPureRef = false;
            var appendNode = PowExpression(skipRangeUnion).Node;
            currentNode = _factory.BinaryNode(la == FormulaLexer.MULT ? '*' : '/', currentNode, appendNode);
        }

        return (isPureRef, currentNode);
    }


    private (bool IsPureRef, TNode Node) PowExpression(bool skipRangeUnion)
    {
        var (isPureRef, currentNode) = PercentExpression(skipRangeUnion);
        for (var la = _la; la == FormulaLexer.POW; la = _la)
        {
            Consume();
            isPureRef = false;
            var appendNode = PercentExpression(skipRangeUnion).Node;
            currentNode = _factory.BinaryNode('^', currentNode, appendNode);
        }

        return (isPureRef, currentNode);
    }

    private (bool IsPureRef, TNode Node) PercentExpression(bool skipRangeUnion)
    {
        var (isPureRef, prefixAtomNode) = PrefixAtomExpression(skipRangeUnion);
        if (_la == FormulaLexer.PERCENT)
        {
            Consume();
            return (false, _factory.Unary('%', prefixAtomNode));
        }

        return (isPureRef, prefixAtomNode);
    }

    private (bool IsPureRef, TNode Node) PrefixAtomExpression(bool skipRangeUnion)
    {
        var la = _la;
        switch (la)
        {
            case FormulaLexer.PLUS:
                Consume();
                var neutralAtom = PrefixAtomExpression(skipRangeUnion).Node;
                return (false, _factory.Unary('+', neutralAtom));

            case FormulaLexer.MINUS:
                Consume();
                var minusAtom = PrefixAtomExpression(skipRangeUnion).Node;
                return (false, _factory.Unary('-', minusAtom));

            default:
                return AtomExpression(skipRangeUnion);
        }
    }

    private (bool IsPureRef, TNode Node) AtomExpression(bool skipRangeUnion)
    {
        var la = _la;
        switch (la)
        {
            // Constant
            case FormulaLexer.NONREF_ERRORS:
            case FormulaLexer.LOGICAL_CONSTANT:
            case FormulaLexer.NUMERICAL_CONSTANT:
            case FormulaLexer.STRING_CONSTANT:
            case FormulaLexer.OPEN_CURLY:
                var constant = Constant();
                return (false, constant);

            // '(' expression ')'
            case FormulaLexer.OPEN_BRACE:
                Consume();
                var (isPureRef, expression) = Expression(false);
                Match(FormulaLexer.CLOSE_BRACE);
                if (isPureRef)
                {
                    if (skipRangeUnion)
                        return (true, RefIntersectionExpression(true, expression));
                    else
                        // The only remaining stuff is for this, isn't recursive
                        return (true, RefExpression(true, expression));
                }

                return (isPureRef, expression);

            // function_call
            case FormulaLexer.CELL_FUNCTION_LIST:
            case FormulaLexer.USER_DEFINED_FUNCTION_NAME:
                var funName = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                var args = ArgumentList();
                return (false, _factory.Function(funName, args));

            // ref_expression
            default:
                if (skipRangeUnion)
                    return (true, RefIntersectionExpression());
                else
                    // The only remaining stuff is for this, isn't recursive
                    return (true, RefExpression());
        }
    }

    private TNode RefExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var currentNode = RefIntersectionExpression(replaceFirstAtom, refAtom);
        for (var la = _la; la == FormulaLexer.COMMA; la = _la)
        {
            Consume();
            var appendNode = RefIntersectionExpression();
            currentNode = _factory.BinaryNode(',', currentNode, appendNode);
        }

        return currentNode;
    }

    private TNode RefIntersectionExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var currentNode = RefRangeExpression(replaceFirstAtom, refAtom);
        for (var la = _la; la == FormulaLexer.SPACE; la = _la)
        {
            Consume();
            var appendNode = RefRangeExpression();
            currentNode = _factory.BinaryNode(' ', currentNode, appendNode);
        }

        return currentNode;
    }

    private TNode RefRangeExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        var currentNode = RefAtomExpression(replaceFirstAtom, refAtom);
        for (var la = _la; la == FormulaLexer.COLON; la = _la)
        {
            Consume();
            var appendNode = RefAtomExpression();
            currentNode = _factory.BinaryNode(':', currentNode, appendNode);
        }

        return currentNode;
    }

    private TNode RefAtomExpression(bool replaceFirstAtom = false, TNode? refAtom = null)
    {
        if (replaceFirstAtom)
            return refAtom!;

        var la = _la;
        switch (la)
        {
            case FormulaLexer.REF_CONSTANT:
                var refError = _factory.ErrorNode(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                return refError;

            case FormulaLexer.OPEN_BRACE:
                Consume();
                var refExpression = RefExpression();
                Match(FormulaLexer.CLOSE_BRACE);
                return refExpression;

            // cell_reference has been inlined

            // local_cell_reference
            case FormulaLexer.A1_REFERENCE:
                var localCellReferenceNode = _factory.LocalCellReference(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                return localCellReferenceNode;

            // external_cell_reference
            case FormulaLexer.BANG_REFERENCE:
            case FormulaLexer.SINGLE_SHEET_REFERENCE:
                var externalCellReferenceNode = _factory.ExternalCellReference(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                return externalCellReferenceNode;

            // external_cell_reference
            case FormulaLexer.SHEET_RANGE_PREFIX:
                var start = _tokenSource.TokenStartCharIndex;
                Consume();
                Match(FormulaLexer.A1_REFERENCE);
                var sheetRangeReferenceNode = _factory.ExternalCellReference(_input, start, _tokenSource.CharIndex - start);
                return sheetRangeReferenceNode;

            // ref_function_call
            case FormulaLexer.REF_FUNCTION_LIST:
                var refFunName = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                var args = ArgumentList();
                return _factory.Function(refFunName, args);

            // name_reference | structure_reference
            case FormulaLexer.NAME:
                var localName = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                if (_la == FormulaLexer.INTRA_TABLE_REFERENCE)
                {
                    var tableReference = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                    Consume();
                    return _factory.StructureReference(localName, tableReference);
                }

                return _factory.LocalNameReference(localName);

            case FormulaLexer.BOOK_PREFIX:
                var bookPrefix = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                var externalName = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Match(FormulaLexer.NAME);
                if (_la == FormulaLexer.INTRA_TABLE_REFERENCE)
                {
                    var intraTableReference = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                    Consume();
                    return _factory.StructureReference(bookPrefix, externalName, intraTableReference);
                }

                return _factory.ExternalNameReference(bookPrefix, externalName);

            case FormulaLexer.SINGLE_SHEET_PREFIX:
                var sheetPrefix = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                if (_la != FormulaLexer.NAME)
                    throw new Exception("Expecred NAME token");

                var name = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                return _factory.LocalNameReference(sheetPrefix, name);

            case FormulaLexer.INTRA_TABLE_REFERENCE:
                // This can be done only if formula is directly in the table
                var localTableReference = _input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                Consume();
                return _factory.StructureReference(localTableReference);
        }

        throw MakeException("Unex[ected token");

    }

    private TNode Constant()
    {
        var la = _la;
        switch (la)
        {
            case FormulaLexer.NONREF_ERRORS:
                var errorNode = _factory.ErrorNode(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - 1);
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
                var array = ConstantListRows();
                Match(FormulaLexer.CLOSE_CURLY);
                return _factory.ArrayNode(array);

            default:
                throw MakeException("Unexpected token");
        }
    }

    private TScalarValue[,] ConstantListRows()
    {
        var elements = new List<TScalarValue>();
        var firstRowWidth = ConstantListRow(elements);
        var height = 1;
        var la = _la;
        while (la == FormulaLexer.SEMICOLON)
        {
            Consume();
            var nextRowWidth = ConstantListRow(elements);
            if (nextRowWidth != firstRowWidth)
                throw new Exception("Sucks");

            height++;
            la = _la;
        }

        var idx = 0;
        var array = new TScalarValue[height, firstRowWidth];
        for (var row = 0; row < height; ++row)
        {
            for (var col = 0; col < firstRowWidth; ++col)
            {
                array[row, col] = elements[idx++];
            }
        }

        return array;
    }

    private int ConstantListRow(List<TScalarValue> buffer)
    {
        var firstValue = ArrayConstant();
        var origSize = buffer.Count;
        buffer.Add(firstValue);
        var la = _la;
        while (la == FormulaLexer.COMMA)
        {
            Consume();
            var next = ArrayConstant();
            buffer.Add(next);
            la = _la;
        }

        return buffer.Count - origSize;
    }

    private TScalarValue ArrayConstant()
    {
        var la = _la;
        TScalarValue? value;
        switch (la)
        {
            case FormulaLexer.REF_CONSTANT:
            case FormulaLexer.NONREF_ERRORS:
                value = _factory.ErrorValue(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                break;
            case FormulaLexer.LOGICAL_CONSTANT:
                value = _factory.LogicalValue(_input[_tokenSource.TokenStartCharIndex] == 'T');
                break;
            case FormulaLexer.MINUS:
                Consume();
                if (_la != FormulaLexer.NUMERICAL_CONSTANT)
                    throw new Exception("Expecting number");
                value = _factory.NumberValue(-ParseNumber(_input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex)));
                break;
            case FormulaLexer.NUMERICAL_CONSTANT:
                value = _factory.NumberValue(ParseNumber(_input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex)));
                break;
            case FormulaLexer.STRING_CONSTANT:
                value = _factory.TextValue(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
                break;
            default:
                throw MakeException("");
        }
        Consume();
        return value;
    }

    private List<TNode> ArgumentList()
    {
        var args = new List<TNode>();

        while (true)
        {
            var la = _la;
            if (la == FormulaLexer.CLOSE_BRACE)
            {
                Consume();
                return args;
            }

            if (la == FormulaLexer.COMMA)
            {
                Consume();
                args.Add(_factory.BlankNode());
            }
            else
            {
                var (_, arg) = Expression(true);
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
            throw UnexpectedToken(expected);

        Consume();
    }

    private void Consume()
    {
        var token = _tokenSource.NextToken();
        _la = token.Type;
    }

    private double ParseNumber(ReadOnlySpan<char> number)
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
        return _factory.LogicalNode(_input[_tokenSource.TokenStartCharIndex] == 'T');
    }

    private TNode ConvertNumber()
    {
        var number = ParseNumber(_input.AsSpan(_tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex));
        var numberNode = _factory.NumberNode(number);
        return numberNode;
    }

    private TNode ConvertText()
    {
        var textNode = _factory.TextNode(_input, _tokenSource.TokenStartCharIndex, _tokenSource.CharIndex - _tokenSource.TokenStartCharIndex);
        return textNode;
    }
    
    private Exception UnexpectedToken(params int[] expectedToken)
    {
        return MakeException($"Unexpected token {GetLaTokenName()}, expected one of {string.Join(",", expectedToken.Select(GetTokenName))}.");
    }

    private Exception MakeException(string message)
    {
        return new Exception($"Error at line {_tokenSource.Line}:{_tokenSource.Column} of '{_input}': {message}");
    }

    private static string GetTokenName(int tokenType) => FormulaLexer.ruleNames[tokenType - 1];

    private string GetLaTokenName() => GetTokenName(_la);
}
