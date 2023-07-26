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
        AppendA1Reference(sb);

        if (RowAbs)
            sb.Append('$');
        sb.Append(Row);
        return sb.ToString();
    }

    private void AppendA1Reference(StringBuilder sb)
    {
        var columnIndex = Column - 1;
        const int aaIndex = 26 * 26;
        if (columnIndex >= aaIndex)
        {
            var thirdIndex = columnIndex / aaIndex;
            var thirdLetter = (char)('A' + thirdIndex);
            sb.Append(thirdLetter);
            columnIndex -= thirdLetter * aaIndex;
        }

        const int aIndex = 26;
        if (columnIndex >= aIndex)
        {
            var secondIndex = columnIndex / aIndex;
            var secondLetter = (char)('A' + secondIndex);
            sb.Append(secondLetter);
            columnIndex -= secondLetter * aIndex;
        }

        sb.Append((char)('A' + columnIndex));
    }
}