using System;

namespace ClosedXML.Parser;

internal sealed class TransformContext
{
    internal TransformContext(string formula, int row, int col)
    {
        if (row is < 1 or > RowCol.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(row));

        if (col is < 1 or > RowCol.MaxCol)
            throw new ArgumentOutOfRangeException(nameof(row));

        Formula = formula;
        Row = row;
        Col = col;
    }

    internal string Formula { get; }

    /// <summary>
    /// Absolute row number in a sheet.
    /// </summary>
    internal int Row { get; }

    /// <summary>
    /// Absolute column number in a sheet.
    /// </summary>
    internal int Col { get; }
}