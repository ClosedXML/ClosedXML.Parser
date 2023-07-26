using System.Text;

namespace ClosedXML.Parser;

public readonly struct CellArea
{
    /// <summary>
    /// Start sheet for references to another sheet or 3D references
    /// </summary>
    public readonly string? StartSheet;

    /// <summary>
    /// End sheet for 3D references. If null, then it is only reference to a worksheet.
    /// </summary>
    public readonly string? EndSheet;

    /// <summary>
    /// Top left corner of a reference.
    /// </summary>
    public readonly CellReference First;

    /// <summary>
    /// Right bottom corner of a reference. If reference is a single cell, the value is same as the <see cref="First"/>.
    /// </summary>
    public readonly CellReference Last;

    /// <summary>
    /// Ctor for 3D references.
    /// </summary>
    public CellArea(string startSheet, string endSheet, CellReference first, CellReference last)
    {
        StartSheet = startSheet;
        EndSheet = endSheet;
        First = first;
        Last = last;
    }

    /// <summary>
    /// Ctor for worksheet references.
    /// </summary>
    public CellArea(string sheet, CellReference first, CellReference last)
    {
        StartSheet = sheet;
        EndSheet = null;
        First = first;
        Last = last;
    }

    /// <summary>
    /// Ctor for current worksheet references.
    /// </summary>
    public CellArea(CellReference first, CellReference last)
    {
        StartSheet = null;
        EndSheet = null;
        First = first;
        Last = last;
    }

    /// <summary>
    /// Ctor for a reference to a single cell.
    /// </summary>
    public CellArea(CellReference cell) : this(cell, cell)
    {
    }

    /// <summary>
    /// Ctor for a reference to a single cell.
    /// </summary>
    public CellArea(int column, int row) : this(new CellReference(false, column, false, row))
    {
    }

    public string GetDisplayString()
    {
        var sb = new StringBuilder();
        if (StartSheet is not null)
            sb.Append(StartSheet);
        if (EndSheet is not null)
        {
            sb.Append(':');
            sb.Append(EndSheet);
        }

        if (sb.Length > 0)
            sb.Append('!');

        sb.Append(First.GetDisplayString());
        if (!First.Equals(Last))
        {
            sb.Append(':');
            sb.Append(Last.GetDisplayString());
        }

        return sb.ToString();
    }
}