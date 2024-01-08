using System;
using JetBrains.Annotations;

namespace ClosedXML.Parser;

/// <summary>
/// Convert between <em>A1</em> and <em>R1C1</em> style formulas.
/// </summary>
[PublicAPI]
public static class FormulaConverter
{
    private static readonly TextVisitorR1C1 s_visitorR1C1 = new();
    private static readonly TextVisitorA1 s_visitorA1 = new();

    /// <summary>
    /// Convert a formula in <em>A1</em> form to the <em>R1C1</em> form.
    /// </summary>
    /// <param name="formulaA1">Formula text.</param>
    /// <param name="row">The row origin of R1C1, from 1 to 1048576.</param>
    /// <param name="col">The column origin of R1C1, from 1 to 16384.</param>
    /// <returns>Formula converted to R1C1.</returns>
    /// <exception cref="ParsingException">The formula is not parseable.</exception>
    public static string ToR1C1(string formulaA1, int row, int col)
    {
        var ctx = new TransformContext(formulaA1, row, col);
        var transformedFormula = FormulaParser<TransformedSymbol, TransformedSymbol, TransformContext>.CellFormulaA1(formulaA1, ctx, s_visitorR1C1);
        return Normalize(transformedFormula, formulaA1);
    }

    /// <summary>
    /// Convert a formula in <em>R1C1</em> form to the <em>A1</em> form.
    /// </summary>
    /// <param name="formulaR1C1">Formula text in R1C1.</param>
    /// <param name="row">The row origin of R1C1, from 1 to 1048576.</param>
    /// <param name="col">The column origin of R1C1, from 1 to 16384.</param>
    /// <returns>Formula converted to A1.</returns>
    /// <exception cref="ParsingException">The formula is not parseable.</exception>
    public static string ToA1(string formulaR1C1, int row, int col)
    {
        var ctx = new TransformContext(formulaR1C1, row, col);
        var transformedFormula = FormulaParser<TransformedSymbol, TransformedSymbol, TransformContext>.CellFormulaR1C1(formulaR1C1, ctx, s_visitorA1);
        return Normalize(transformedFormula, formulaR1C1);
    }

    private static string Normalize(TransformedSymbol transformedFormula, string originalFormula)
    {
        // Because of intersection operator, we trim the whitespaces at the end before sending
        // formula to the parser. Add them back, if necessary.
        var trimmed = originalFormula.TrimEnd();
        var endLength = originalFormula.Length - trimmed.Length;
        var trimmedEnd = originalFormula.AsSpan().Slice(trimmed.Length, endLength);
        return transformedFormula.ToString(trimmedEnd);
    }

    private class TextVisitorR1C1 : FormulaGeneratorVisitor
    {
        protected override ReferenceArea ModifyRef(TransformContext ctx, ReferenceArea reference)
        {
            return reference.ToR1C1(ctx.Row, ctx.Col);
        }

        protected override RowCol ModifyCellFunction(TransformContext ctx, RowCol cell)
        {
            return cell.ToR1C1(ctx.Row, ctx.Col);
        }
    }

    private class TextVisitorA1 : FormulaGeneratorVisitor
    {
        protected override ReferenceArea ModifyRef(TransformContext ctx, ReferenceArea reference)
        {
            return reference.ToA1(ctx.Row, ctx.Col);
        }

        protected override RowCol ModifyCellFunction(TransformContext ctx, RowCol cell)
        {
            return cell.ToA1(ctx.Row, ctx.Col);
        }
    }
}