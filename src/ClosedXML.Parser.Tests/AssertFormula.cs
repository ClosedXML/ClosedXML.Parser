using Antlr4.Runtime;
using Antlr4.Runtime.Atn;

namespace ClosedXML.Parser.Tests;

internal static class AssertFormula
{
    public static void CstParsed(string formula)
    {
        var inputStream = new AntlrInputStream(formula);
        var lexer = new FormulaLexer(inputStream);
        var commonTokenStream = new CommonTokenStream(lexer);
        var parser = new FormulaParser(commonTokenStream, TextWriter.Null, TextWriter.Null)
        {
            Interpreter =  {  PredictionMode = PredictionMode.SLL  }
        };
        var res = parser.formula();
        
        Assert.IsNull(res.exception);
    }
}