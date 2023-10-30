namespace ClosedXML.Parser;

/// <summary>
/// Convert between <em>A1</em> and <em>R1C1</em> style formulas.
/// </summary>
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
        return FormulaParser<string, string, (int Row, int Col)>.CellFormulaA1(formulaA1, (row, col), s_visitorR1C1);
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
        return FormulaParser<string, string, (int Row, int Col)>.CellFormulaR1C1(formulaR1C1, (row, col), s_visitorA1);
    }

    private class TextVisitorR1C1 : FormulaGeneratorVisitor
    {
        protected override ReferenceSymbol ModifyRef(ReferenceSymbol reference, (int Row, int Col) point)
        {
            return reference.ToR1C1(point.Row, point.Col);
        }

        protected override RowCol ModifyCellFunction(RowCol cell, (int Row, int Col) point)
        {
            return cell.ToR1C1(point.Row, point.Col);
        }
    }

    private class TextVisitorA1 : FormulaGeneratorVisitor
    {
        protected override ReferenceSymbol ModifyRef(ReferenceSymbol reference, (int Row, int Col) point)
        {
            return reference.ToA1(point.Row, point.Col);
        }

        protected override RowCol ModifyCellFunction(RowCol cell, (int Row, int Col) point)
        {
            return cell.ToA1(point.Row, point.Col);
        }
    }
}