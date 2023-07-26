namespace ClosedXML.Parser;

internal static class StructuredReferenceAreaExtensions
{
    public static string GetDisplayString(this StructuredReferenceArea area)
    {
        return area switch
        {
            StructuredReferenceArea.None => String.Empty,
            StructuredReferenceArea.Data => "[#Data]",
            StructuredReferenceArea.Headers => "[#Headers]",
            StructuredReferenceArea.Totals => "[#Totals]",
            StructuredReferenceArea.Data | StructuredReferenceArea.Headers => "[#Headers], [#Data]",
            StructuredReferenceArea.Data | StructuredReferenceArea.Totals => "[#Data], [#Totals]",
            StructuredReferenceArea.All => "[#All]",
            StructuredReferenceArea.ThisRow => "[#This Row]",
            _ => throw new NotSupportedException($"Unexpected area {area}.")
        };
    }
}