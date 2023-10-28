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
/// <see cref="ReferenceSymbol"/> reference style, e.g. column 14 with
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

    /// <summary>
    /// Create a new <see cref="RowCol"/> with both row and columns specified.
    /// </summary>
    /// <param name="rowType">The type used to interpret the row position.</param>
    /// <param name="rowValue">The value for the row position.</param>
    /// <param name="columnType">The type used to interpret the column position.</param>
    /// <param name="columnValue">The value for the column position.</param>
    internal RowCol(ReferenceAxisType rowType, int rowValue, ReferenceAxisType columnType, int columnValue)
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

    /// <summary>
    /// Create a new <see cref="RowCol"/> with both row and columns specified.
    /// </summary>
    /// <param name="rowAbs">Is the row reference absolute? If false, then relative.</param>
    /// <param name="rowValue">The value for the row position.</param>
    /// <param name="colAbs">Is the column reference absolute? If false, then relative.</param>
    /// <param name="columnValue">The value for the column position.</param>
    internal RowCol(bool rowAbs, int rowValue, bool colAbs, int columnValue)
        : this(rowAbs ? Absolute : Relative, rowValue, colAbs ? Absolute : Relative, columnValue)
    {
    }

    /// <summary>
    /// Create a new <see cref="RowCol"/> with both row and columns specified
    /// with relative values. Used mostly for A1 style.
    /// </summary>
    /// <param name="row">The relative position of the row.</param>
    /// <param name="column">The relative position of the column.</param>
    internal RowCol(int row, int column)
        : this(Relative, row, Relative, column)
    {
    }

    /// <summary>
    /// Compares two <see cref="RowCol"/> objects by value. The result specifies whether
    /// all properties of the two <see cref="RowCol"/> objects are equal.
    /// </summary>
    public static bool operator ==(RowCol lhs, RowCol rhs) => lhs.Equals(rhs);

    /// <summary>
    /// Compares two <see cref="RowCol"/> objects by value. The result specifies whether
    /// any property of the two <see cref="RowCol"/> objects is not equal.
    /// </summary>
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

    /// <summary>
    /// Get the representation of the <see cref="RowCol"/> as a text in R1C1
    /// style.
    /// </summary>
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

    /// <summary>
    /// Convert A1 RowCol to a R1C1 string representation.
    /// </summary>
    /// <remarks>There is no check that RowCol was parsed through A1.</remarks>
    /// <param name="sb">String buffer to write the representation.</param>
    /// <param name="row">Actual row of a cell.</param>
    /// <param name="col">Actual column of a cell.</param>
    internal void ToR1C1(StringBuilder sb, int row, int col)
    {
        // TODO: Is this stupid idea? Maybe I should convert RowCol from R1C1 and then use GetDisplayStringR1C1.
        AppendAxis(sb, 'R', RowType, RowValue, row);
        AppendAxis(sb, 'C', ColumnType, ColumnValue, col);

        static void AppendAxis(StringBuilder sb, char axisName, ReferenceAxisType axisType, int axisValue, int actual)
        {
            // None is ignored because that means other axis is full row/column.
            if (axisType == Relative)
            {
                sb.Append(axisName);
                if (axisValue != actual)
                    sb.Append('[').Append(axisValue - actual).Append(']');
            }
            else if (axisType == Absolute)
            {
                sb.Append(axisName).Append(axisValue);
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

    /// <summary>
    /// Check whether the <paramref name="obj"/> is of type <see cref="RowCol"/>
    /// and all values are same as this one.
    /// </summary>
    public override bool Equals(object obj)
    {
        return obj is RowCol other && Equals(other);
    }

    /// <summary>
    /// Check whether the all values of <paramref name="other"/> are same as
    /// this one.
    /// </summary>
    public bool Equals(RowCol other)
    {
        return ColumnType == other.ColumnType &&
               ColumnValue == other.ColumnValue &&
               RowType == other.RowType &&
               RowValue == other.RowValue;
    }

    /// <summary>
    /// Returns a hash code for this <see cref="RowCol"/>.
    /// </summary>
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