using System.Text;

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

    public string GetDisplayString()
    {
        var sb = new StringBuilder();
        if (ColumnAbs)
            sb.Append('$');
        sb.Append(GetA1Reference());

        if (RowAbs)
            sb.Append('$');
        sb.Append(Row);
        return sb.ToString();
    }

    private string GetA1Reference()
    {
        var columnIndex = Column;
        var column = string.Empty;
        do
        {
            columnIndex -= 1;
            var index = columnIndex % 26;
            columnIndex -= index;
            columnIndex /= 26;
            column = (char)('A' + index) + column;
        } while (columnIndex > 0);

        return column;
    }
}