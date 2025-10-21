using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace ClosedXML.Parser.Pratt;

/// <summary>
/// A lexer for pratt parser.
/// </summary>
internal class Lexer
{
    private const string OPERATOR_CHARS = "!,;^*/+-&=<>%:#@(){} ";
    private const int EOF = -1;

    private static readonly bool[] IsOp;

    private readonly Queue<Token> _queue = new(4);
    private string _input = string.Empty; // Currently tokenized formula
    private int _start; // The start index of currently parsed token in Next()
    private int _i; // Index of current code point _c in _input
    private int _c; // A current code point (including astral planes) or -1 if at the EOF


    static Lexer()
    {
        IsOp = new bool[128];
        foreach (var op in OPERATOR_CHARS)
            IsOp[op] = true;
    }

    public Lexer()
        : this(string.Empty)
    {
    }

    public Lexer(string input)
    {
        Reset(input);
    }

    private bool IsEof => _c == EOF;

    /// <summary>
    /// Prepare lexer to start tokenization of the <paramref name="formula"/>.
    /// </summary>
    /// <param name="formula">Formula to tokenize.</param>
    public void Reset(string formula)
    {
        _input = formula ?? throw new ArgumentNullException();
        _start = -1;
        _i = -1;
        _c = 0;
    }

    public Token Consume()
    {
        if (_queue.Count == 0)
            return Next();

        return _queue.Dequeue();
    }

    public Token Peek(int distance = 1)
    {
        // TODO: Replace BCL queue with a structure that allows index access
        while (_queue.Count < distance)
            _queue.Enqueue(Next());

        var enumerator = _queue.GetEnumerator();
        for (var i = 0; i < distance; ++i)
            enumerator.MoveNext();

        return enumerator.Current;
    }

    private Token Next()
    {
        if (_i < 0)
            Advance();

        if (IsEof)
            return new Token(TokenType.Eof, 0, 0);

        _start = _i;

        // Number
        if (IsDigit(_c))
        {
            // Whole number part
            DigitSequence();

            // Fractional part
            if (_c == '.')
            {
                Advance();
                DigitSequence();
            }

            ExponentPart();

            return T(TokenType.Number);
        }
        if (_c is '.')
        {
            Advance();

            // Fractional part
            DigitSequence();
            ExponentPart();

            return T(TokenType.Number);
        }

        // Text
        if (_c == '"')
        {
            Advance();

            while (!IsEof)
            {
                if (_c == '"')
                {
                    Advance();
                    if (_c != '"')
                        return T(TokenType.Text);
                }

                if (!IsXml10Char(_c))
                    throw new ParsingException($"Invalid text character (codepoint {_c:x8}).");

                Advance();
            }

            throw ParsingException.UnterminatedLiteral(_start, '"');
        }

        // QIdent
        if (_c == '\'')
        {
            while (!IsEof)
            {
                Advance();

                if (_c == '\'')
                {
                    Advance();
                    if (_c != '\'')
                        return T(TokenType.QIdent);
                }
            }

            throw ParsingException.UnterminatedLiteral(_start, '\'');
        }

        if (IsIdentStart(_c))
        {
            Advance();
            while (!IsEof && IsIdentNext(_c))
                Advance();

            return T(TokenType.Ident);
        }

        if (_c == '+')
            return FoundToken(TokenType.Plus);

        if (_c == '-')
            return FoundToken(TokenType.Minus);

        if (_c == '*')
            return FoundToken(TokenType.Mul);

        if (_c == '/')
            return FoundToken(TokenType.Div);

        if (_c == '^')
            return FoundToken(TokenType.Pow);

        if (_c == '%')
            return FoundToken(TokenType.Percent);

        if (_c == '&')
            return FoundToken(TokenType.Concat);

        if (_c == '!')
            return FoundToken(TokenType.Bang);

        if (_c == '(')
            return FoundToken(TokenType.LeftParen);

        if (_c == ')')
            return FoundToken(TokenType.RightParen);

        if (_c == '{')
            return FoundToken(TokenType.LeftCurly);

        if (_c == '}')
            return FoundToken(TokenType.RightCurly);

        if (_c == ',')
            return FoundToken(TokenType.Comma);

        if (_c == ';')
            return FoundToken(TokenType.Semicolon);

        if (_c == ':')
            return FoundToken(TokenType.Range);

        if (_c == '@')
            return FoundToken(TokenType.Intersection);

        // Comparison
        if (_c == '=')
            return FoundToken(TokenType.Equal);

        if (_c == '<')
        {
            var next = Advance();
            if (next == '>')
                return FoundToken(TokenType.NotEqual);

            if (next == '=')
                return FoundToken(TokenType.LessEqual);

            return T(TokenType.Less);
        }

        if (_c == '>')
        {
            if (Advance() == '=')
            {
                return FoundToken(TokenType.GreaterEqual);
            }

            return T(TokenType.Greater);
        }

        if (IsWhitespace(_c))
        {
            do
            {
                Advance();
            } while (IsWhitespace(_c));

            return T(TokenType.Whitespace);
        }

        // Spill operator and errors
        if (_c == '#')
        {
            var char1 = Advance();
            switch (char1)
            {
                case 'D' or 'd':
                    return Error("#DIV/0!", 2);
                case 'R' or 'r':
                    return Error("#REF!", 2);
                case 'V' or 'v':
                    return Error("#VALUE!", 2);
                case 'G' or 'g':
                    return Error("#GETTING_DATA", 2);
                case 'N' or 'n':
                    {
                        var char2 = Advance();
                        if (char2 == '/')
                            return Error("#N/A", 3);

                        if (char2 == 'A')
                            return Error("#NAME?", 3);

                        var char3 = Advance();
                        if (char2 == 'U' && char3 == 'L')
                            return Error("#NULL!", 4);

                        if (char2 == 'U' && char3 == 'M')
                            return Error("#NUM!", 4);

                        throw ParsingException.TokenPartialMatch(_start, TokenType.Error);
                    }
            }

            return T(TokenType.Spill);
        }

        if (_c == '[')
        {
            var level = 0;
            do
            {
                switch (_c)
                {
                    case '[':
                        ++level;
                        break;
                    case ']':
                        --level;
                        break;
                    case '\'':
                        Advance(); // Escaped chars don't change level - skip
                        break;
                }

                if (IsEof)
                    throw new ParsingException($"Unable to find closing square bracket for token from position {_start}.");

                if (level > 2)
                    throw new ParsingException($"There can be at most two nested square brackets in a token from position {_start}.");

                Advance();
            } while (level > 0);

            return T(TokenType.SquareIdent);
        }

        throw ParsingException.UnableToSelectToken(_start);

        static bool IsWhitespace(int c)
        {
            return c is ' ' or '\r' or '\n' or '\t';
        }

        static bool IsDigit(int c)
        {
            return c is >= '0' and <= '9';
        }

        // Check [0-9]+
        void DigitSequence()
        {
            do
            {
                if (!IsDigit(_c))
                    throw ParsingException.TokenPartialMatch(_start, TokenType.Number);
                Advance();
            }
            while (!IsEof && IsDigit(_c));
        }

        void ExponentPart()
        {
            if (_c is 'e' or 'E')
            {
                if (Advance() is '+' or '-')
                    Advance();

                DigitSequence();
            }
        }

        static bool IsIdentStart(int c)
        {
            // Ident must satisfy logical-literal, sheet-name, name and A1-cell/column/row
            return
                IsAsciiLetter(c) || // name + A1
                c == '$' || // A1
                (c is '_' or '\\' or '?') || // name
                (c > 0x7F && IsLetterOrLetterMark(c)); // name
        }

        static bool IsIdentNext(int c)
        {
            // Stop at operators
            if (c < IsOp.Length && IsOp[c])
                return false;

            return IsIdentStart(c) ||
                c is >= '0' and <= '9' ||  // name, A1
                c == '.'; // name + future-functions
        }

        Token Error(string error, int start)
        {
            foreach (var errorChar in error.AsSpan().Slice(start))
            {
                Advance();
                if (ToUpperAlpha(_c) != errorChar)
                    throw ParsingException.TokenPartialMatch(_start, TokenType.Error);
            }

            Advance();
            return T(TokenType.Error);
        }

        // Token that ends at the Current has been found. Advance to next and return token.
        Token FoundToken(TokenType type)
        {
            Advance();
            return T(type);
        }

        // Convert a-z to A-Z, keep other codepoints same.
        static int ToUpperAlpha(int codepoint)
        {
            return codepoint is >= 'a' and <= 'z'
                ? 'A' + codepoint - 'a'
                : codepoint;
        }

        static bool IsAsciiLetter(int codepoint)
        {
            // Convert to lowercase, normalize 'a' to 0, and check if within 0 (~A)..25(~Z).
            // Really cool use of cast int to uint (-1 to 0xFFFFFFFF), thus saving one comparison
            // and avoiding potential pipeline stall.
            return (uint)((codepoint | 32) - 97) <= 25U;
        }

        static bool IsLetterOrLetterMark(int codepoint)
        {
            // TODO: Only netstandard 2.1 has a parameter of type int, 2.0 has only char.
            if (codepoint > 0xFFFF)
                return false; // No letters from astral planes for us :(

            // Letters are categories from 0 to OtherLetter category. Then there are NonSpacingMark (accents and such).
            return CharUnicodeInfo.GetUnicodeCategory((char)codepoint) <= UnicodeCategory.NonSpacingMark;
        }

        // Is codepoint a character per XML 1.0 spec (2.2)?
        static bool IsXml10Char(int codepoint)
        {
            // .NET is using a lookup table with properties
            if (codepoint <= 0xFFFF)
                return XmlConvert.IsXmlChar((char)codepoint);

            return codepoint <= 0x10FFFF;
        }
    }

    private Token T(TokenType type)
    {
        return new Token(type, _start, _i);
    }

    private int Advance()
    {
        if (++_i >= _input.Length)
        {
            _c = -1;
            return (char)_c;
        }

        var c = _input[_i];

        if (char.IsLowSurrogate(c))
            throw ParsingException.UnpairedSurrogate(c, _i);

        if (char.IsHighSurrogate(c))
        {
            if (_i + 1 >= _input.Length)
                throw ParsingException.UnpairedSurrogate(c, _i);

            var low = _input[++_i];
            if (!char.IsLowSurrogate(low))
                throw ParsingException.UnpairedSurrogate(c, _i - 1);

            _c = char.ConvertToUtf32(c, low);
            return _c;
        }

        return _c = c;
    }
}
