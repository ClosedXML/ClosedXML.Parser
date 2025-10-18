using System;
using ClosedXML.Parser.Pratt;

namespace ClosedXML.Parser;

/// <summary>
/// Indicates an error during parsing. In most cases, unexpected token.
/// </summary>
public class ParsingException : Exception
{
    internal ParsingException(string message) : base(message)
    {
    }

    /// <summary>
    /// There are problems with underlying stream.
    /// </summary>
    internal static ParsingException UnpairedSurrogate(int codepoint, int position)
    {
        throw new ParsingException($"Found an unpaired surrogate 0x{codepoint:X4} at {position}.");
    }

    /// <summary>
    /// Token has a start and end indicator, but no end indicator was found.
    /// </summary>
    internal static Exception UnterminatedLiteral(int start, char delimiter)
    {
        throw new ParsingException($"An unterminated literal (delimiter {delimiter}) found at position {start}.");
    }

    /// <summary>
    /// A token was started to be parsed, but is not complete or there is a problem with it.
    /// </summary>
    internal static Exception TokenPartialMatch(int start, TokenType type)
    {
        throw new ParsingException($"Token {type} was parsed from position {start}, but was only partially matched.");
    }

    /// <summary>
    /// Lexer has no idea which token it should start to parse.
    /// </summary>
    internal static Exception UnableToSelectToken(int start)
    {
        throw new ParsingException($"Unable to determine a token at position {start}.");
    }
}
