namespace ClosedXML.Parser;

public readonly struct CellArea
{
    public readonly CellReference First;
    public readonly CellReference Last;

    public CellArea(CellReference first, CellReference last)
    {
        First = first;
        Last = last;
    }

    public CellArea(CellReference cell) : this(cell, cell)
    {
    }

    public CellArea(int column, int row) : this(new CellReference(false, column, false, row))
    {
    }
}