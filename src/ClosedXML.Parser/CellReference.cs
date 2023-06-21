namespace ClosedXML.Parser;

public readonly struct CellReference
{
    public readonly bool ColumnAbs;
    public readonly int Column;
    public readonly bool RowAbs;
    public readonly int Row;

    public CellReference(bool colAbs, int col, bool rowAbs, int row)
    {
        ColumnAbs = colAbs;
        Column = col;
        RowAbs = rowAbs;
        Row = row;
    }

    public CellReference(int column, int row) : this(false, column, false, row)
    {
    }
}