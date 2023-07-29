namespace ClosedXML.Parser.Tests.Rules;

public class FormulaRuleTests
{
    [Fact]
    public void Additional_text_after_expression_is_error()
    {
        AssertFormula.CheckParsingErrorContains("A1)", "The formula `A1)` wasn't parsed correctly. The expression `A1` was parsed, but the rest `)` wasn't.");
    }
}