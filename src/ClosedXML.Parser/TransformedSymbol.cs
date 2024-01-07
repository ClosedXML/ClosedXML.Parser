using System;

namespace ClosedXML.Parser;

/// <summary>
/// A symbol that represents a transformed symbol value. Should be used during AST transformation.
/// </summary>
internal readonly struct TransformedSymbol
{
    private readonly string _formulaText;

    /// <summary>
    /// Range of the symbol in original formula.
    /// </summary>
    private readonly SymbolRange _range;

    /// <summary>
    /// The text that replaced symbol or null, if there was no change.
    /// </summary>
    private readonly string? _transformedText;

    private TransformedSymbol(string formulaText, string? transformedText, SymbolRange range) : this()
    {
        _formulaText = formulaText;
        _transformedText = transformedText;
        _range = range;
    }

    /// <summary>
    /// Length of the transformed symbol.
    /// </summary>
    public int Length => _transformedText?.Length ?? _range.End - _range.Start;

    /// <summary>
    /// There was no transformation of the symbol.
    /// </summary>
    /// <param name="formula">Text of whole formula.</param>
    /// <param name="range">Range of the symbol in the formula.</param>
    /// <param name="transformedSymbol">The string of a transformed symbol.</param>
    public static TransformedSymbol ToText(string formula, SymbolRange range, string transformedSymbol)
    {
        return new TransformedSymbol(formula, transformedSymbol, range);
    }

    /// <summary>
    /// There was no transformation of the symbol.
    /// </summary>
    /// <param name="formula">Text of whole formula.</param>
    /// <param name="range">Range of the symbol in the formula.</param>
    public static TransformedSymbol CopyOriginal(string formula, SymbolRange range)
    {
        return new TransformedSymbol(formula, null, range);
    }

    public ReadOnlySpan<char> AsSpan()
    {
        if (_transformedText is not null)
            return _transformedText.AsSpan();

        return _formulaText.AsSpan(_range.Start, _range.End - _range.Start);
    }
};