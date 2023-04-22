namespace ClosedXML.Parser.Tests;

[TestClass]
public class ConstantParsingTests
{
    [TestMethod]
    [DataRow("1")]
    [DataRow("10.5")]
    [DataRow("10.5E5")]
    [DataRow(".1E-4")]
    public void Number_is_parsed(string formula)
    {
        AssertFormula.CstParsed(formula);
    }

    [TestMethod]
    [DataRow("#REF!")]
    [DataRow("#VALUE!")]
    public void Error_is_parsed(string formula)
    {
        AssertFormula.CstParsed(formula);
    }
}