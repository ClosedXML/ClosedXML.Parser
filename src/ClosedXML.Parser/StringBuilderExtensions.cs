using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// Extension methods for building formulas.
/// </summary>
internal static class StringBuilderExtensions
{
    public static StringBuilder AppendSheetName(this StringBuilder sb, string sheetName)
    {
        return NameUtils.EscapeName(sb, sheetName);
    }

    public static StringBuilder AppendExternalSheetName(this StringBuilder sb, int workbookIndex, string sheetName)
    {
        if (NameUtils.ShouldQuote(sheetName.AsSpan()))
        {
            return sb
                .Append('\'')
                .AppendBookIndex(workbookIndex)
                .AppendEscapedSheetName(sheetName)
                .Append('\'');
        }

        return sb
            .AppendBookIndex(workbookIndex)
            .AppendSheetName(sheetName);
    }
    public static StringBuilder AppendEscapedSheetName(this StringBuilder sb, string sheetName)
    {
        var startIndex = sb.Length;
        return sb.Append(sheetName).Replace("'", "''", startIndex, sheetName.Length);
    }

    public static StringBuilder AppendReferenceSeparator(this StringBuilder sb)
    {
        return sb.Append('!');
    }

    public static StringBuilder AppendBookIndex(this StringBuilder sb, int bookIndex)
    {
        return sb.Append('[').Append(bookIndex).Append(']');
    }

    public static StringBuilder AppendFunction(this StringBuilder sb, TransformContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        return sb.Append(functionName).AppendArguments(ctx, range, arguments);
    }

    public static StringBuilder AppendRef(this StringBuilder sb, ReferenceArea? reference)
    {
        return reference is null ? sb.Append("#REF!") : reference.Value.Append(sb);
    }

    public static StringBuilder AppendRef(this StringBuilder sb, RowCol? rowCol)
    {
        if (rowCol is null)
            return sb.Append("#REF!");

        rowCol.Value.Append(sb);
        return sb;
    }

    public static StringBuilder AppendStartFragment(this StringBuilder sb, TransformContext ctx, SymbolRange symbolRange, TransformedSymbol nestedNode)
    {
        var formula = ctx.Formula;
        for (var i = symbolRange.Start; i < nestedNode.OriginalRange.Start; ++i)
            sb.Append(formula[i]);

        return sb;
    }

    public static StringBuilder AppendMiddleFragment(this StringBuilder sb, TransformContext ctx, TransformedSymbol beforeNode, TransformedSymbol afterNode)
    {
        var formula = ctx.Formula;
        for (var i = beforeNode.OriginalRange.End; i < afterNode.OriginalRange.Start; ++i)
            sb.Append(formula[i]);

        return sb;
    }

    public static StringBuilder AppendEndFragment(this StringBuilder sb, TransformContext ctx, SymbolRange symbolRange, TransformedSymbol nestedNode)
    {
        var formula = ctx.Formula;
        for (var i = nestedNode.OriginalRange.End; i < symbolRange.End; ++i)
            sb.Append(formula[i]);

        return sb;
    }

    public static StringBuilder AppendArguments(this StringBuilder sb, TransformContext ctx, SymbolRange range, IReadOnlyList<TransformedSymbol> arguments)
    {
        if (arguments.Count == 0)
        {
            var braceIdx = GetStartBraceIndex(ctx, range, range.End);
            var braces = ctx.Formula.AsSpan().Slice(braceIdx, range.End - braceIdx);
            sb.Append(braces);
        }
        else
        {
            sb
                .AppendStartBrace(ctx, range, arguments[0])
                .AppendArguments(ctx, arguments)
                .AppendEndFragment(ctx, range, arguments[arguments.Count - 1]);
        }

        return sb;
    }

    private static StringBuilder AppendStartBrace(this StringBuilder sb, TransformContext ctx, SymbolRange range, TransformedSymbol firstNode)
    {
        var firstNodeStart = firstNode.OriginalRange.Start;
        var braceIdx = GetStartBraceIndex(ctx, range, firstNodeStart);
        for (var j = braceIdx; j < firstNodeStart; ++j)
            sb.Append(ctx.Formula[j]);

        return sb;
    }

    private static int GetStartBraceIndex(TransformContext ctx, SymbolRange range, int nodeStart)
    {
        var formula = ctx.Formula;
        var braceIdx = nodeStart - 1;
        for (; braceIdx > range.Start; --braceIdx)
        {
            if (formula[braceIdx] == '(')
                return braceIdx;
        }

        throw new InvalidOperationException("No opening brace found.");
    }

    private static StringBuilder AppendArguments(this StringBuilder sb, TransformContext ctx, IReadOnlyList<TransformedSymbol> arguments)
    {
        if (arguments.Count > 0)
            sb.Append(arguments[0].AsSpan());

        for (var i = 1; i < arguments.Count; ++i)
        {
            sb.AppendMiddleFragment(ctx, arguments[i - 1], arguments[i]);
            sb.Append(arguments[i].AsSpan());
        }

        return sb;
    }

#if NETSTANDARD2_0
    /// <summary>
    /// Compatibility method for NETStandard 2.0, which doesn't have methods with <c>Span</c> arguments.
    /// </summary>
    public static StringBuilder Append(this StringBuilder sb, ReadOnlySpan<char> span)
    {
        foreach (var c in span)
            sb.Append(c);

        return sb;
    }
#endif
}