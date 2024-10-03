using ClosedXML.Parser.Rolex;
using JetBrains.Annotations;
using System;
using System.Collections.Generic;

namespace ClosedXML.Parser;

/// <summary>
/// A utility class that parses various types of references.
/// </summary>
public static class ReferenceParser
{
    /// <summary>
    /// Parses area reference in A1 form. The possibilities are
    /// <list type="bullet">
    ///   <item>Cell (e.g. <c>F8</c>).</item>
    ///   <item>Area (e.g. <c>B2:$D7</c>).</item>
    ///   <item>Colspan (e.g. <c>$D:$G</c>).</item>
    ///   <item>Rowspan (e.g. <c>14:$15</c>).</item>
    /// </list>
    /// Doesn't allow any whitespaces or extra values inside.
    /// </summary>
    /// <param name="text">Text to parse.</param>
    /// <param name="area">Parsed area.</param>
    /// <returns><c>true</c> if parsing was a success, <c>false</c> otherwise.</returns>
    [PublicAPI]
    public static bool TryParseA1(string text, out ReferenceArea area)
    {
        if (text is null)
            throw new ArgumentNullException();

        // a1_reference : A1_CELL
        //              | A1_CELL COLON A1_CELL
        //              | A1_SPAN_REFERENCE
        var tokens = RolexLexer.GetTokensA1(text.AsSpan());
        var isValid = IsA1Reference(tokens);
        if (!isValid)
        {
            area = default;
            return false;
        }

        area = TokenParser.ParseReference(text.AsSpan(), isA1: true);
        return true;
    }

    /// <summary>
    /// Parses area reference in A1 form. The possibilities are
    /// <list type="bullet">
    ///   <item>Cell (e.g. <c>F8</c>).</item>
    ///   <item>Area (e.g. <c>B2:$D7</c>).</item>
    ///   <item>Colspan (e.g. <c>$D:$G</c>).</item>
    ///   <item>Rowspan (e.g. <c>14:$15</c>).</item>
    /// </list>
    /// Doesn't allow any whitespaces or extra values inside.
    /// </summary>
    /// <exception cref="ParsingException">Invalid input.</exception>
    [PublicAPI]
    public static ReferenceArea ParseA1(string text)
    {
        if (!TryParseA1(text, out var area))
            throw new ParsingException($"Unable to parse '{text}'.");

        return area;
    }

    /// <summary>
    /// Try to parse a A1 reference that has a sheet (e.g. <c>'Data values'!A$1:F10</c>).
    /// If <paramref name="text"/> contains only reference without a sheet or anything
    /// else (e.g. <c>A1</c>), return <c>false</c>.
    /// </summary>
    /// <remarks>
    /// The method doesn't accept
    /// <list type="bullet">
    ///   <item>Sheet names, e.g. <c>Sheet!name</c>.</item>
    ///   <item>External sheet references, e.g. <c>[1]Sheet!A1</c>.</item>
    ///   <item>Sheet errors, e.g. <c>Sheet5!$REF!</c>.</item>
    /// </list>
    /// </remarks>
    /// <param name="text">Text to parse.</param>
    /// <param name="sheetName">Name of the sheet, unescaped (e.g. the sheetName will contain <c>Jane's</c> for <c>'Jane''s'!A1</c>).</param>
    /// <param name="area">Parsed reference.</param>
    /// <returns><c>true</c> if parsing was a success, <c>false</c> otherwise.</returns>
    [PublicAPI]
    public static bool TryParseSheetA1(string text, out string sheetName, out ReferenceArea area)
    {
        if (text is null)
            throw new ArgumentNullException(nameof(text));

        var tokens = RolexLexer.GetTokensA1(text.AsSpan());
        if (tokens.Count == 0 ||
            tokens[0].SymbolId != Token.SINGLE_SHEET_PREFIX)
        {
            sheetName = string.Empty;
            area = default;
            return false;
        }

        var sheetPrefixToken = tokens[0];
        var sheetPrefix = text.AsSpan(sheetPrefixToken.StartIndex, sheetPrefixToken.Length);
        TokenParser.ParseSingleSheetPrefix(sheetPrefix, out int? workbookIndex, out sheetName);
        if (workbookIndex is not null)
        {
            sheetName = string.Empty;
            area = default;
            return false;
        }

        tokens.RemoveAt(0);
        if (!IsA1Reference(tokens))
        {
            sheetName = string.Empty;
            area = default;
            return false;
        }

        var referenceArea = text.AsSpan().Slice(sheetPrefixToken.Length);
        area = TokenParser.ParseReference(referenceArea, isA1: true);
        return true;
    }

    private static bool IsA1Reference(IReadOnlyList<Token> tokens)
    {
        var isValid = tokens.Count switch
        {
            2 => tokens[0].SymbolId is Token.A1_CELL or Token.A1_SPAN_REFERENCE &&
                 tokens[1].SymbolId == Token.EofSymbolId,
            4 => tokens[0].SymbolId == Token.A1_CELL &&
                 tokens[1].SymbolId == Token.COLON &&
                 tokens[2].SymbolId == Token.A1_CELL &&
                 tokens[3].SymbolId == Token.EofSymbolId,
            _ => false,
        };
        return isValid;
    }
}
