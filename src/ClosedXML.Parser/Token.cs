using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

// ReSharper disable InconsistentNaming
namespace ClosedXML.Parser;

/// <summary>
/// A token for a formula input.
/// </summary>
internal readonly struct Token
{
    private static readonly IReadOnlyDictionary<int, string> SymbolNames = typeof(Token)
        .GetFields(BindingFlags.Public | BindingFlags.Static)
        .Where(f => f.FieldType == typeof(int) && f.IsLiteral)
        .ToDictionary(x => (int)x.GetValue(null), x => x.Name);

    /// <summary>
    /// An error symbol id.
    /// </summary>
    public const int ErrorSymbolId = -2;

    /// <summary>
    /// An symbol id for end of file. Mostly for compatibility with ANTLR.
    /// </summary>
    public const int EofSymbolId = -1;

    public const int REF_CONSTANT = 1;
    public const int NONREF_ERRORS = 2;
    public const int LOGICAL_CONSTANT = 3;
    public const int NUMERICAL_CONSTANT = 4;
    public const int STRING_CONSTANT = 5;
    public const int POW = 6;
    public const int MULT = 7;
    public const int DIV = 8;
    public const int PLUS = 9;
    public const int MINUS = 10;
    public const int CONCAT = 11;
    public const int EQUAL = 12;
    public const int NOT_EQUAL = 13;
    public const int LESS_OR_EQUAL_THAN = 14;
    public const int LESS_THAN = 15;
    public const int GREATER_OR_EQUAL_THAN = 16;
    public const int GREATER_THAN = 17;
    public const int PERCENT = 18;
    public const int SEMICOLON = 19;
    public const int COLON = 20;
    public const int OPEN_BRACE = 21;
    public const int CLOSE_BRACE = 22;
    public const int OPEN_CURLY = 23;
    public const int CLOSE_CURLY = 24;
    public const int COMMA = 25;
    public const int SPACE = 26;
    public const int INTERSECT = 27;
    public const int SPILL = 28;
    public const int BOOK_PREFIX = 29;
    public const int BANG_REFERENCE = 30;
    public const int SHEET_RANGE_PREFIX = 31;
    public const int SINGLE_SHEET_PREFIX = 32;
    public const int A1_REFERENCE = 33;
    public const int REF_FUNCTION_LIST = 34;
    public const int CELL_FUNCTION_LIST = 35;
    public const int USER_DEFINED_FUNCTION_NAME = 36;
    public const int NAME = 37;
    public const int INTRA_TABLE_REFERENCE = 38;

    /// <summary>
    /// A token ID or TokenType. Non-negative integer. The values are from Antlr grammar, starting with 1.
    /// See <c>FormulaLexer.tokens</c>. The value -1 indicates an error and unrecognized token and is always
    /// last token.
    /// </summary>
    public readonly int SymbolId;

    /// <summary>
    /// The starting index of a token, in code units (=chars).
    /// </summary>
    public readonly int StartIndex;

    /// <summary>
    /// Length of a token in code units (=chars). For non-error tokens, must be at least 1. Ignore for error token.
    /// </summary>
    public readonly int Length;

    public Token(int symbolId, int startIndex, int length)
    {
        SymbolId = symbolId;
        StartIndex = startIndex;
        Length = length;
    }

    public static Token EofSymbol(int index) => new(EofSymbolId, index, 0);

    public static string GetSymbolName(int symbolId)
    {
        if (!SymbolNames.TryGetValue(symbolId, out var name))
            throw new ArgumentOutOfRangeException($"Invalid symbol {symbolId}.");

        return name;
    }

    public bool Equals(Token other)
    {
        return SymbolId == other.SymbolId && StartIndex == other.StartIndex && Length == other.Length;
    }

    public override bool Equals(object? obj)
    {
        return obj is Token other && Equals(other);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = SymbolId;
            hashCode = (hashCode * 397) ^ StartIndex;
            hashCode = (hashCode * 397) ^ Length;
            return hashCode;
        }
    }

    public override string ToString() => $"Symbol: {SymbolId}; StartIdx: {StartIndex}; Len: {Length}";
}