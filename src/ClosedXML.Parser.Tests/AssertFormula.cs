using Antlr4.Runtime;
using Antlr4.Runtime.Atn;

namespace ClosedXML.Parser.Tests;

internal static class AssertFormula
{
    public static void CstParsed(string formula)
    {
        var inputStream = new AntlrInputStream(formula);
        var lexer = new FormulaLexer(inputStream);
        var listener = new LexerErrorListener();
        lexer.RemoveErrorListeners();
        lexer.AddErrorListener(listener);
        var commonTokenStream = new CommonTokenStream(lexer);
        var parser = new FormulaParser(commonTokenStream, TextWriter.Null, TextWriter.Null)
        {
            Interpreter = { PredictionMode = PredictionMode.SLL }
        };
        parser.ErrorListeners.Clear();

        var res = parser.formula();

        Assert.IsNull(listener.Errors, $"{formula}  {listener.Errors}");
        Assert.IsNull(res.exception, $"{formula} {res.exception}");
    }

    private class LexerErrorListener : IAntlrErrorListener<int>
    {
        internal string? Errors { get; private set; }

        public void SyntaxError(TextWriter output, IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            var error = $"line {line}:{charPositionInLine} {msg}";
            Errors = Errors is null ? error : Errors + "\n" + error;
        }
    }
}