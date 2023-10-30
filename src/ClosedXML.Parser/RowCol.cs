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
    public ReferenceAxisType ColumnType { get; }

    /// <summary>
    /// Position of a column.
    /// </summary>
    public int ColumnValue { get; }

    /// <summary>
    /// How to interpret the <see cref="RowValue"/> value.
    /// </summary>
    public ReferenceAxisType RowType { get; }

    /// <summary>
    /// Position of a row.
    /// </summary>
    public int RowValue { get; }

    /// <summary>
    /// Does <c>RowCol</c> use <em>A1</em> semantic?
    /// </summary>
    public bool IsA1 => Style == A1;

    /// <summary>
    /// Does <c>RowCol</c> use <em>R1C1</em> semantic?
    /// </summary>
    public bool IsR1C1 => Style == R1C1;

    /// <summary>
    /// Reference style of the <c>RowCol</c>.
    /// </summary>
    public ReferenceStyle Style { get; }

    /// <summary>
    /// Create a new <see cref="RowCol"/> with both row and columns specified.
    /// </summary>
    /// <param name="rowType">The type used to interpret the row position.</param>
    /// <param name="rowValue">The value for the row position.</param>
    /// <param name="columnType">The type used to interpret the column position.</param>
    /// <param name="columnValue">The value for the column position.</param>
    /// <param name="style">Semantic of the reference.</param>
    internal RowCol(ReferenceAxisType rowType, int rowValue, ReferenceAxisType columnType, int columnValue, ReferenceStyle style)
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
        Style = style;
    }

    /// <summary>
    /// Create a new <see cref="RowCol"/> with both row and columns specified.
    /// </summary>
    /// <param name="rowAbs">Is the row reference absolute? If false, then relative.</param>
    /// <param name="rowValue">The value for the row position.</param>
    /// <param name="colAbs">Is the column reference absolute? If false, then relative.</param>
    /// <param name="columnValue">The value for the column position.</param>
    /// <param name="style">Semantic of the reference.</param>
    internal RowCol(bool rowAbs, int rowValue, bool colAbs, int columnValue, ReferenceStyle style)
        : this(rowAbs ? Absolute : Relative, rowValue, colAbs ? Absolute : Relative, columnValue, style)
    {
    }

    /// <summary>
    /// Create a new <see cref="RowCol"/> with both row and columns specified
    /// with relative values. Used mostly for A1 style.
    /// </summary>
    /// <param name="row">The relative position of the row.</param>
    /// <param name="column">The relative position of the column.</param>
    /// <param name="style">Semantic of the reference.</param>
    internal RowCol(int row, int column, ReferenceStyle style)
        : this(Relative, row, Relative, column, style)
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
    /// Get a reference in <em>A1</em> notation.
    /// </summary>
    /// <exception cref="InvalidOperationException">When <c>RowCol</c> doesn't use <em>A1</em> semantic.</exception>
    public string GetDisplayStringA1()
    {
        if (!IsA1)
            throw new InvalidOperationException("RowCol doesn't use A1 semantic.");

        var sb = new StringBuilder();
        AppendA1(sb);
        return sb.ToString();
    }

    /// <summary>
    /// Get a reference in <em>R1C1</em> notation.
    /// </summary>
    /// <exception cref="InvalidOperationException">When <c>RowCol</c> doesn't use <em>R1C1</em> semantic.</exception>
    public string GetDisplayStringR1C1()
    {
        if (!IsR1C1)
            throw new InvalidOperationException("RowCol doesn't use R1C1 semantic.");

        var sb = new StringBuilder();
        AppendR1C1(sb);
        return sb.ToString();
    }

    /// <summary>
    /// Convert <c>RowCol</c> to <em>R1C1</em>.
    /// </summary>
    /// <remarks>If <c>RowCol</c> already is in <em>R1C1</em>, return it directly.</remarks>
    /// <param name="anchorRow">A row coordinate that should be used as an anchor for relative R1C1 reference.</param>
    /// <param name="anchorCol">A column coordinate that should be used as an anchor for relative R1C1 reference.</param>
    /// <returns>RowCol with R1C1 semantic.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Row or col is out of valid row or column number.</exception>
    public RowCol ToR1C1(int anchorRow, int anchorCol)
    {
        if (anchorRow is < 1 or > 1048576)
            throw new ArgumentOutOfRangeException(nameof(anchorRow));

        if (anchorCol is < 1 or > 16384)
            throw new ArgumentOutOfRangeException(nameof(anchorCol));

        if (IsR1C1)
            return this;

        var newRowPosition = ConvertAxis(RowType, RowValue, anchorRow);
        var newColPosition = ConvertAxis(ColumnType, ColumnValue, anchorCol);

        return new RowCol(RowType, newRowPosition, ColumnType, newColPosition, R1C1);

        static int ConvertAxis(ReferenceAxisType axisType, int axisValue, int anchorPosition)
        {
            return axisType switch
            {
                Relative => axisValue - anchorPosition,
                Absolute => axisValue,
                None => 0,
                _ => throw new NotSupportedException()
            };
        }
    }

    /// <summary>
    /// Convert <c>RowCol</c> to <em>A1</em>.
    /// </summary>
    /// <remarks>If <c>RowCol</c> already is in <em>A1</em>, return it directly.</remarks>
    /// <param name="anchorRow">A row coordinate that should be used as an anchor for relative <em>R1C1</em> reference.</param>
    /// <param name="anchorCol">A column coordinate that should be used as an anchor for relative <em>R1C1</em> reference.</param>
    /// <returns>RowCol with R1C1 semantic.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Row or col is out of valid row or column number.</exception>
    public RowCol ToA1(int anchorRow, int anchorCol)
    {
        if (anchorRow is < 1 or > 1048576)
            throw new ArgumentOutOfRangeException(nameof(anchorRow));

        if (anchorCol is < 1 or > 16384)
            throw new ArgumentOutOfRangeException(nameof(anchorCol));

        if (IsA1)
            return this;

        var newRowPosition = ConvertAxis(RowType, RowValue, anchorRow);
        var newColPosition = ConvertAxis(ColumnType, ColumnValue, anchorCol);

        return new RowCol(RowType, newRowPosition, ColumnType, newColPosition, A1);

        static int ConvertAxis(ReferenceAxisType axisType, int axisValue, int anchorPosition)
        {
            return axisType switch
            {
                Relative => axisValue + anchorPosition,
                Absolute => axisValue,
                None => 0,
                _ => throw new NotSupportedException()
            };
        }
    }

    /// <inheritdoc cref="GetDisplayStringA1()"/>
    /// <param name="sb">String buffer where to write the output.</param>
    /// <exception cref="InvalidOperationException">When <c>RowCol</c> is not in <em>A1</em> notation.</exception>
    internal void AppendA1(StringBuilder sb)
    {
        if (!IsA1)
            throw new InvalidOperationException("RowCol doesn't use A1 semantic.");

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
    }

    /// <inheritdoc cref="GetDisplayStringR1C1()"/>
    /// <param name="sb">String buffer where to write the output.</param>
    /// <exception cref="InvalidOperationException">When <c>RowCol</c> is not in <em>A1</em> notation.</exception>
    internal void AppendR1C1(StringBuilder sb)
    {
        if (!IsR1C1)
            throw new InvalidOperationException("RowCol doesn't use R1C1 semantic.");

        AppendAxis(sb, 'R', RowType, RowValue);
        AppendAxis(sb, 'C', ColumnType, ColumnValue);

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
               RowValue == other.RowValue &&
               IsA1 == other.IsA1;
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
            hashCode = (hashCode * 397) ^ IsA1.GetHashCode();
            return hashCode;
        }
    }
}