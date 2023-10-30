namespace ClosedXML.Parser;

internal static class ReferenceAreaExtensions
{
    public static string GetDisplayString(this ReferenceArea reference, ReferenceStyle style)
    {
        return style switch
        {
            ReferenceStyle.A1 => reference.GetDisplayStringA1(),
            ReferenceStyle.R1C1 => reference.GetDisplayStringR1C1(),
            _ => throw new NotSupportedException()
        };
    }
}