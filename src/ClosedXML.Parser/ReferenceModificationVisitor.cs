using System;
using System.Text;

namespace ClosedXML.Parser;

internal class ReferenceModificationVisitor : CopyVisitor
{
    private const string REF_ERROR = "#REF!";

    public override TransformedSymbol ErrorNode(TransformContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        if (range.Length == error.Length)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Deal with `Sheet!REF!`, `#REF!A1` and `#REF!#REF!`
        var symbol = ctx.Formula.AsSpan().Slice(range.Start, range.Length);
        var sheetIsRefError = symbol.StartsWith(REF_ERROR.AsSpan(), StringComparison.OrdinalIgnoreCase);
        var referenceIsRefError = symbol.EndsWith(REF_ERROR.AsSpan(), StringComparison.OrdinalIgnoreCase);

        if (sheetIsRefError && referenceIsRefError)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        if (sheetIsRefError)
        {
            var referenceSymbol = symbol.Slice(REF_ERROR.Length);
            var reference = TokenParser.ParseReference(referenceSymbol, ctx.IsA1);
            var modifiedReference = ModifyRef(ctx, reference);
            var nodeText = new StringBuilder()
                .Append(symbol.Slice(0, REF_ERROR.Length))
                .AppendRef(modifiedReference)
                .ToString();
            return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
        }
        else
        {
            var sheet = symbol.Slice(0, symbol.Length - REF_ERROR.Length - 1);
            var modifiedSheet = ModifySheet(ctx, sheet.ToString());
            var nodeText = new StringBuilder()
                .AppendSheetReference(modifiedSheet)
                .Append(symbol.Slice(symbol.Length - REF_ERROR.Length))
                .ToString();
            return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
        }
    }

    public override TransformedSymbol Reference(TransformContext ctx, SymbolRange range, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return base.Reference(ctx, range, modifiedReference.Value);
    }

    public override TransformedSymbol SheetReference(TransformContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return base.SheetReference(ctx, range, sheet, modifiedReference.Value);
    }

    public override TransformedSymbol Reference3D(TransformContext ctx, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return base.Reference3D(ctx, range, firstSheet, lastSheet, modifiedReference.Value);
    }

    public override TransformedSymbol ExternalSheetReference(TransformContext ctx, SymbolRange range, int workbookIndex,
        string sheet, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return base.ExternalSheetReference(ctx, range, workbookIndex, sheet, modifiedReference.Value);
    }

    public override TransformedSymbol ExternalReference3D(TransformContext ctx, SymbolRange range, int workbookIndex,
        string firstSheet, string lastSheet, ReferenceArea reference)
    {
        var modifiedReference = ModifyRef(ctx, reference);
        if (modifiedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, REF_ERROR);

        return base.ExternalReference3D(ctx, range, workbookIndex, firstSheet, lastSheet, modifiedReference.Value);
    }

    /// <summary>
    /// Modify reference to a cell.
    /// </summary>
    /// <param name="ctx">The origin of formula.</param>
    /// <param name="reference">Area reference.</param>
    /// <returns>Modified reference or null if <c>#REF!</c>.</returns>
    protected virtual ReferenceArea? ModifyRef(TransformContext ctx, ReferenceArea reference)
    {
        return reference;
    }
}