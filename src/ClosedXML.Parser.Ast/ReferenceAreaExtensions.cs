namespace ClosedXML.Parser;

internal static class ReferenceAreaExtensions
{
    public static string GetDisplayString(this ReferenceArea area, ReferenceStyle style)
    {
        return style switch
        {
            ReferenceStyle.A1 => area.GetDisplayStringA1(),
            ReferenceStyle.R1C1 => area.GetDisplayStringR1C1(),
            _ => throw new NotSupportedException()
        };
    }
}