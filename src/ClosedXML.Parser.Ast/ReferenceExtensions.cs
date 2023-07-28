namespace ClosedXML.Parser;

internal static class ReferenceExtensions
{
    public static string GetDisplayString(this Reference reference, ReferenceStyle style)
    {
        return style switch
        {
            ReferenceStyle.A1 => reference.GetDisplayStringA1(),
            ReferenceStyle.R1C1 => throw new NotImplementedException(),
            _ => throw new NotSupportedException()
        };
    }
}