using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// A reference to a sheet defined by row and column axis.
/// </summary>
public readonly struct Reference
{
    /// <summary>
    /// An unspeciied reference. Used for areas, when only one corner is specified.
    /// </summary>
    public static readonly Reference Missing = new(ReferenceAxisType.None, 0, ReferenceAxisType.None, 0);

    /// <summary>
    /// How to interpret the <see cref="ColumnValue"/> value.
    /// </summary>
    public readonly ReferenceAxisType ColumnType;

    /// <summary>
    /// Position of a column.
    /// </summary>
    public readonly int ColumnValue;

    /// <summary>
    /// How to interpret the <see cref="RowValue"/> value.
    /// </summary>
    public readonly ReferenceAxisType RowType;

    /// <summary>
    /// Position of a row.
    /// </summary>
    public readonly int RowValue;

    public Reference(ReferenceAxisType columnType, int columnValue, ReferenceAxisType rowType, int rowValue)
    {
        ColumnType = columnType;
        ColumnValue = columnValue;
        RowType = rowType;
        RowValue = rowValue;
    }

    public Reference(bool colAbs, int columnValue, bool rowAbs, int rowValue)
        : this(colAbs ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, columnValue, rowAbs ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, rowValue)
    {
    }

    public Reference(int column, int row)
        : this(ReferenceAxisType.Relative, column, ReferenceAxisType.Relative, row)
    {
    }

    public string GetDisplayString()
    {
        var sb = new StringBuilder();
        if (ColumnType == ReferenceAxisType.Absolute)
            sb.Append('$');
        sb.Append(GetA1Reference());

        if (RowType == ReferenceAxisType.Absolute)
            sb.Append('$');
        sb.Append(RowValue);
        return sb.ToString();
    }

    private string GetA1Reference()
    {
        var columnIndex = ColumnValue;
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