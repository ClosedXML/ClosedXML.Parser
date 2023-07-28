namespace ClosedXML.Parser;

internal static class ReferenceAreaExtensions
{
    public static string GetDisplayString(this ReferenceArea area, ReferenceStyle style)
    {
        return style switch
        {
            ReferenceStyle.A1 => area.GetDisplayStringA1(),
            ReferenceStyle.R1C1 => throw new NotImplementedException(),
            _ => throw new NotSupportedException()
        };
    }
}