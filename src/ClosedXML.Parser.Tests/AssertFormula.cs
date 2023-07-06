using Antlr4.Runtime;
using Antlr4.Runtime.Atn;

namespace ClosedXML.Parser.Tests;

internal static class AssertFormula
{
    /// <summary>
    /// Assert that a formula is parsed into a single childless node.
    /// </summary>
    public static void SingleNodeParsed<TNode>(string formula, TNode expectedNode)
        where TNode : AstNode
    {
        var lexer = new FormulaLexer(new CodePointCharStream(formula), TextWriter.Null, TextWriter.Null);
        var parser = new FormulaParser<ScalarValue, AstNode>(formula, lexer, new F());
        var node = (TNode)parser.Formula();
        Assert.AreEqual(0, node.Children.Length);
        Assert.AreEqual(expectedNode, node);
    }
    
    /// <summary>
    /// Assert that text is recognized as a single token of a token type.
    /// </summary>
    /// <param name="tokenText">Text that should contain a single token.</param>
    /// <param name="tokenType">Expected token type, from <see cref="FormulaLexer"/> const .</param>
    public static void AssertTokenType(string tokenText, int tokenType)
    {
        var commonTokenStream = new CommonTokenStream(new FormulaLexer(new AntlrInputStream(tokenText)));
        commonTokenStream.Fill();
        Assert.AreEqual(2, commonTokenStream.Size);
        Assert.AreEqual(tokenType, commonTokenStream.Get(0).Type);
        Assert.AreEqual(FormulaLexer.Eof, commonTokenStream.Get(1).Type);
    }

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