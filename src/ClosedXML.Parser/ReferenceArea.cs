using System.Text;

namespace ClosedXML.Parser;

/// <summary>
/// Due to frequency of an area in formulas, the grammar has a token that represents
/// an area in a sheet. This is the DTO from parser to engine. Two corners make an area
/// for A1 notation, but not for R1C1 (has several edge cases).
/// </summary>
public readonly struct ReferenceArea
{
    /// <summary>
    /// First reference. First in terms of position in formula, not position in sheet.
    /// </summary>
    public readonly Reference First;

    /// <summary>
    /// Second reference. Second in terms of position in formula, not position in sheet.
    /// </summary>
    public readonly Reference Second;

    public ReferenceArea(Reference first, Reference second)
    {
        First = first;
        Second = second;
    }

    public string GetDisplayString()
    {
        var sb = new StringBuilder();
        sb.Append(First.GetDisplayString());
        if (First == Second)
            return First.GetDisplayString();

        return new StringBuilder()
            .Append(First.GetDisplayString())
            .Append(':')
            .Append(Second.GetDisplayString())
            .ToString();
    }
}