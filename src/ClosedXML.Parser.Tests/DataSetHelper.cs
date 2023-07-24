using CsvHelper.Configuration.Attributes;
using CsvHelper.Configuration;
using CsvHelper;
using JetBrains.Annotations;
using System.Globalization;

namespace ClosedXML.Parser.Tests;

internal static class DataSets
{
    public static IEnumerable<string> ReadCsv(string filename)
    {
        var config = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = false };
        using var reader = new StreamReader(filename);
        using var csv = new CsvReader(reader, config);
        var formulas = csv.GetRecords<Formula>();
        foreach (var formula in formulas)
            yield return formula.Text;
    }

    [UsedImplicitly]
    private record Formula([Index(0)] string Text);
}