using System;
using System.Text;
using static ClosedXML.Parser.ReferenceAxisType;

namespace ClosedXML.Parser;

/// <summary>
/// <para>
/// One endpoint of a reference defined by row and column axis. It can be
/// <list type="bullet">
///   <item>
///   A single cell that is an intersection of row and a column
///   </item>
///   <item>
///   An entire row, e.g. <c><em>A</em>:B</c> or <c><em>R5</em>:R10</c>.
///   </item>
///   <item>
///   An entire column, e.g. <c><em>7</em>:14</c> or <c><em>C7</em>:C10</c>.
///   </item>
/// </list>
/// The content of values and thus their interpretation depends on the
/// <see cref="ReferenceArea"/> reference style, e.g. column 14 with
/// <see cref="Relative"/> can indicate <c>R[14]</c> or <c>X14</c> for A1
/// style.
/// </para>
/// <para>
/// Not all combinations are valid and the content of the reference corresponds
/// to a valid token in expected reference style (e.g. in R1C1, <c>R</c> is
/// a valid standalone reference, but there is no such possibility for A1).
/// </para>
/// </summary>
public readonly struct RowCol : IEquatable<RowCol>
{
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

    public RowCol(ReferenceAxisType columnType, int columnValue, ReferenceAxisType rowType, int rowValue)
    {
        if (columnType == None && rowType == None)
            throw new ArgumentException("At least one of axis must be non-none.");

        if (columnType == None && columnValue != 0)
            throw new ArgumentException("Value for `None` type must be zero.", nameof(columnValue));

        if (rowType == None && rowValue != 0)
            throw new ArgumentException("Value for `None` type must be zero.", nameof(rowValue));

        ColumnType = columnType;
        ColumnValue = columnValue;
        RowType = rowType;
        RowValue = rowValue;
    }

    public RowCol(bool colAbs, int columnValue, bool rowAbs, int rowValue)
        : this(colAbs ? Absolute : Relative, columnValue, rowAbs ? Absolute : Relative, rowValue)
    {
    }

    public RowCol(int column, int row)
        : this(Relative, column, Relative, row)
    {
    }

    public static bool operator ==(RowCol lhs, RowCol rhs) => lhs.Equals(rhs);

    public static bool operator !=(RowCol lhs, RowCol rhs) => !(lhs == rhs);

    /// <summary>
    /// Get a reference in A1 notation. The content must have been created from A1
    /// token, otherwise the output won't be correct.
    /// </summary>
    public string GetDisplayStringA1()
    {
        var sb = new StringBuilder();
        switch (ColumnType)
        {
            case Absolute:
                sb.Append('$').Append(GetA1Reference());
                break;

            case Relative:
                sb.Append(GetA1Reference());
                break;

            case None:
                break;

            default:
                throw new NotSupportedException();
        }

        switch (RowType)
        {
            case Absolute:
                sb.Append('$').Append(RowValue);
                break;

            case Relative:
                sb.Append(RowValue);
                break;

            case None:
                break;

            default:
                throw new NotSupportedException();
        }

        return sb.ToString();
    }

    public string GetDisplayStringR1C1()
    {
        var sb = new StringBuilder();

        AppendAxis(sb, 'R', RowType, RowValue);
        AppendAxis(sb, 'C', ColumnType, ColumnValue);
        return sb.ToString();

        static void AppendAxis(StringBuilder sb, char axis, ReferenceAxisType type, int position)
        {
            switch (type)
            {
                case Absolute:
                    sb.Append(axis).Append(position);
                    break;

                case Relative when position != 0:
                    sb.Append(axis).Append('[').Append(position).Append(']');
                    break;

                case Relative:
                    // position is always 0
                    sb.Append(axis);
                    break;

                case None:
                    break;

                default:
                    throw new NotSupportedException();
            }
        }
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
        return obj is RowCol other && Equals(other);
    }

    public bool Equals(RowCol other)
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