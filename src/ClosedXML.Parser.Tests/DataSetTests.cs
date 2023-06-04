using System.Diagnostics;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using System.Globalization;
using JetBrains.Annotations;
using Antlr4.Runtime;
using ClosedXML.Parser.Tests.Rds;

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

    [TestMethod]
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
            badFormulas.UnionWith(Read(badFormulaPath));
        
        var sw = Stopwatch.StartNew();
        var formulaCount = 0;
        foreach (var formula in Read(input))
        {
            formulaCount++;
            try
            {
                var stream = new CodePointCharStream(formula);
                var lexer = new FormulaLexer(stream);
                lexer.RemoveErrorListeners();
                var parser = new Lexer.FormulaParser<ScalarValue, AstNode>(formula, lexer, new F());
                parser.Formula();
                Assert.IsFalse(badFormulas.Contains(formula), formula);
            }
            catch (Exception e)
            {
                Assert.IsTrue(badFormulas.Contains(formula), $"Parsing '{formula}' failes: {e.Message}");
            }
        }
        
        sw.Stop();
        Console.WriteLine($"Parsed {formulaCount} formulas in {sw.ElapsedMilliseconds}ms ({sw.ElapsedMilliseconds * 1000d / formulaCount:N3}μs/formula)");
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