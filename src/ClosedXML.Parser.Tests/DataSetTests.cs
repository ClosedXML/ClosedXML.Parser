using System.Diagnostics;
using Xunit.Abstractions;

namespace ClosedXML.Parser.Tests;

public class DataSetTests
{
    private readonly ITestOutputHelper _output;

    public DataSetTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
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

    [Fact]
    public void Euses_data_set_is_parseable()
    {
        Assert_formulas_parsed_or_not_as_expected(
            "./data/euses/formulas.csv",
            new[]
            {
                "./data/euses/invalid-external-cell-reference.csv",
                "./data/euses/known-fails.csv"
            });
    }

    private void Assert_formulas_parsed_or_not_as_expected(string input, string[] badFormulaPaths)
    {
        var badFormulas = new HashSet<string>();
        foreach (var badFormulaPath in badFormulaPaths)
            badFormulas.UnionWith(DataSets.ReadCsv(badFormulaPath));

        // Read to memory before the parsing to measure only parsing.
        var formulas = DataSets.ReadCsv(input).ToList();
        var sw = Stopwatch.StartNew();
        var formulaCount = 0;
        foreach (var formula in formulas)
        {
            formulaCount++;
            try
            {
                _ = FormulaParser<ScalarValue, AstNode>.CellFormulaA1(formula, new F());
                Assert.False(badFormulas.Contains(formula), formula);
            }
            catch (Exception e)
            {
                Assert.True(badFormulas.Contains(formula), $"Parsing formula '{formula}' failed: {e.Message}");
            }
        }

        sw.Stop();
        var averageLength = formulas.Sum(x => x.Length) / (double)formulas.Count;
        _output.WriteLine($"Parsed {formulaCount} formulas (Average length {averageLength:F1}) in {sw.ElapsedMilliseconds}ms ({sw.ElapsedMilliseconds * 1000d / formulaCount:N3}μs/formula).");
    }
}