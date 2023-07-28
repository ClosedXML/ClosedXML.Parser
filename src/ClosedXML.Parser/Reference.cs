using System;
using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// A reference to a sheet defined by row and column axis. Not all combinations are
/// valid and the content of the reference corresponds to a valid token in expected
/// reference style (e.g. in R1C1, <c>R</c> is a valid standalone reference, but
/// there is no such possibility for A1).
/// </summary>
public readonly struct Reference : IEquatable<Reference>
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

    public static bool operator ==(Reference lhs, Reference rhs) => lhs.Equals(rhs);

    public static bool operator !=(Reference lhs, Reference rhs) => !(lhs == rhs);

    /// <summary>
    /// Get a reference in A1 notation. The content must have been created from A1
    /// token, otherwise the output won't be correct.
    /// </summary>
    public string GetDisplayStringA1()
    {
        var sb = new StringBuilder();
        switch (ColumnType)
        {
            case ReferenceAxisType.Absolute:
                sb.Append('$').Append(GetA1Reference());
                break;

            case ReferenceAxisType.Relative:
                sb.Append(GetA1Reference());
                break;

            case ReferenceAxisType.None:
                break;

            default:
                throw new NotSupportedException();
        }

        switch (RowType)
        {
            case ReferenceAxisType.Absolute:
                sb.Append('$').Append(RowValue);
                break;

            case ReferenceAxisType.Relative:
                sb.Append(RowValue);
                break;

            case ReferenceAxisType.None:
                break;

            default:
                throw new NotSupportedException();
        }

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

    public override bool Equals(object obj)
    {
        return obj is Reference other && Equals(other);
    }

    public bool Equals(Reference other)
    {
        return ColumnType == other.ColumnType &&
               ColumnValue == other.ColumnValue &&
               RowType == other.RowType &&
               RowValue == other.RowValue;
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = (int)ColumnType;
            hashCode = (hashCode * 397) ^ ColumnValue;
            hashCode = (hashCode * 397) ^ (int)RowType;
            hashCode = (hashCode * 397) ^ RowValue;
            return hashCode;
        }
    }

}