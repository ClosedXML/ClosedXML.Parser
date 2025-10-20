using System.Diagnostics;
using ClosedXML.Parser.Pratt;

namespace ClosedXML.Parser.Tests.Lexers;

public class PrattParserPrecedenceTests
{
    [Theory]
    [InlineData("1+2+3+4", "(((1+2)+3)+4)")]
    [InlineData("1-2-3-4", "(((1-2)-3)-4)")]
    [InlineData("1-2+3-4+5", "((((1-2)+3)-4)+5)")]
    [InlineData("1*2*3*4", "(((1*2)*3)*4)")]
    [InlineData("1/2/3/4", "(((1/2)/3)/4)")]
    [InlineData("1*2/3*4/5", "((((1*2)/3)*4)/5)")]
    [InlineData("2^3^4^5", "(((2^3)^4)^5)")] // Even exponential is left-associative in Excel, contrary to standard convention
    public void Operations_with_same_precedence_are_left_associative(string formula, string normalizedForm)
    {
        AssertSameFormulas(formula, normalizedForm);
    }

    [Theory]
    [InlineData("1+(2+3+4)+((5+6)+7)", "((1+((2+3)+4))+((5+6)+7))")]
    [InlineData("1-(2-3-4)-((5-6)-7)", "((1-((2-3)-4))-((5-6)-7))")]
    [InlineData("1-(2+3-4)+((5-6)+7)", "((1-((2+3)-4))+((5-6)+7))")]
    [InlineData("1*(2*3*4)*((5*6)*7)", "((1*((2*3)*4))*((5*6)*7))")]
    [InlineData("1/(2/3/4)/((5/6)/7)", "((1/((2/3)/4))/((5/6)/7))")]
    [InlineData("1/(2*3/4)*((5/6)*7)", "((1/((2*3)/4))*((5/6)*7))")]
    [InlineData("2^(3^4)^5", "((2^(3^4))^5)")]
    public void Groups_override_precedence(string formula, string normalizedForm)
    {
        AssertSameFormulas(formula, normalizedForm);
    }

    [Theory]
    [InlineData("1+2*3+4/5*6^7-8", "(((1+(2*3))+((4/5)*(6^7)))-8)")]
    [InlineData("1+2-3*4+5/6^7-8*9", "((((1+2)-(3*4))+(5/(6^7)))-(8*9))")]
    public void Operations_are_grouped_by_precedence(string formula, string normalizedForm)
    {
        AssertSameFormulas(formula, normalizedForm);
    }

    private static void AssertSameFormulas(string formula, string normalizedForm)
    {
        var parser = ParserFactory.Create(new F());
        var root = parser.ParseFormula(formula, new Ctx());

        Assert.Equal(normalizedForm, GetNormalizedForm(root));
    }

    private static string GetNormalizedForm(AstNode node)
    {
        return node switch
        {
            ValueNode value => value.GetDisplayString(A1),
            BinaryNode binaryOp => "(" + 
                                   GetNormalizedForm(binaryOp.Children[0]) + 
                                   binaryOp.GetDisplayString(A1) +
                                   GetNormalizedForm(binaryOp.Children[1]) + 
                                   ")",
            _ => throw new UnreachableException()
        };
    }
}
