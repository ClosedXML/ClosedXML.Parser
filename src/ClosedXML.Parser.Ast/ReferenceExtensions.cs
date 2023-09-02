namespace ClosedXML.Parser;

internal static class ReferenceExtensions
{
    public static string GetDisplayString(this RowCol rowCol, ReferenceStyle style)
    {
        return style switch
        {
            ReferenceStyle.A1 => rowCol.GetDisplayStringA1(),
            ReferenceStyle.R1C1 => rowCol.GetDisplayStringR1C1(),
            _ => throw new NotSupportedException()
        };
    }
}