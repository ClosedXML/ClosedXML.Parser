using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using System.Globalization;
using JetBrains.Annotations;

namespace ClosedXML.Parser.Tests;

[TestClass]
public class DataSetTests
{
    [TestMethod]
    public void Enron_data_set_is_parseable()
    {
        Assert_formulas_parsed_or_not_as_expected(
            "./data/enron/formulas.csv",
            new[]
            {
                "./data/enron/invalid-external-cell-reference.csv",
                "./data/enron/known-fails.csv"
            });
    }

    private void Assert_formulas_parsed_or_not_as_expected(string input, string[] badFormulaPaths)
    {
        var badFormulas = new HashSet<string>();
        foreach (var badFormulaPath in badFormulaPaths)
            badFormulas.UnionWith(Read(badFormulaPath));

        foreach (var formula in Read(input))
        {
            try
            {
                AssertFormula.CstParsed(formula);
            }
            catch (Exception)
            {
                Assert.IsTrue(badFormulas.Contains(formula), formula);
            }
        }
    }

    private IEnumerable<string> Read(string filename)
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