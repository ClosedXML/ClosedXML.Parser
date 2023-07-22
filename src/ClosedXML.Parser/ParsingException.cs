using System;

namespace ClosedXML.Parser;

/// <summary>
/// Indicates an error during parsing. In most cases, unexpected token.
/// </summary>
public class ParsingException : Exception
{
    public ParsingException(string message) : base(message)
    {
    }
}