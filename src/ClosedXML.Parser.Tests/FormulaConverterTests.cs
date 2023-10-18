namespace ClosedXML.Parser.Tests;

public class FormulaConverterTests
{
    [Theory]
    [InlineData("TRUE+FALSE", "TRUE+FALSE")] // LogicalNode
    [InlineData("14.75", "14.75")] // NumberNode
    [InlineData("\"Sally's \"\"Emporium\"\"\"", "\"Sally's \"\"Emporium\"\"\"")] // TextNode
    [InlineData("\"A\" & \"B\"", "\"A\"&\"B\"")] // BinaryNode for concat
    [InlineData("1+2", "1+2")] // BinaryNode for plus
    public void Non_reference_tokens_are_not_changed(string formulaA1, string formulaR1C1)
    {
        Assert.Equal(formulaR1C1, FormulaConverter.ToR1C1(formulaA1, 1, 1));
    }
}