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
        var parser = new FormulaParser<ScalarValue, AstNode>(formula, new F());
        var node = (TNode)parser.Formula();
        Assert.Equal(expectedNode, node);
    }

    public static void CheckParsingErrorContains(string formula, string errorSubstring)
    {
        var parser = new FormulaParser<ScalarValue, AstNode>(formula, new F());
        var ex = Assert.Throws<ParsingException>(() => parser.Formula());
        Assert.True(ex.Message.Contains(errorSubstring), $"Error message '{ex.Message}' doesn't contain '{errorSubstring}'."); ;
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
        Assert.Equal(2, commonTokenStream.Size);
        Assert.Equal(tokenType, commonTokenStream.Get(0).Type);
        Assert.Equal(FormulaLexer.Eof, commonTokenStream.Get(1).Type);
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

        Assert.True(listener.ErrorStartIndex is null, $"{formula}  {listener.ErrorStartIndex}");
        Assert.True(res.exception is null, $"{formula} {res.exception}");
    }

    /// <summary>
    /// Get tokens from ANTLR lexer. If there is an error in the <paramref name="formula"/>, insert error token.
    /// </summary>
    public static IReadOnlyList<Token> GetAntlrTokens(string formula)
    {
        var lexer = new FormulaLexer(new CodePointCharStream(formula), TextWriter.Null, TextWriter.Null);
        var listener = new LexerErrorListener();
        lexer.RemoveErrorListeners();
        lexer.AddErrorListener(listener);
        var tokens = lexer.GetAllTokens().Select(x => new Token(x.Type, x.StartIndex, x.StopIndex - x.StartIndex + 1)).ToList();
        if (listener.ErrorStartIndex is not null)
        {
            // Lexer tries to recover. That is good in most cases, but in our case, it's not very
            // compatible with Rolex lexer. Remove the tokens after recovery.
            var errorStartIndex = listener.ErrorStartIndex.Value;
            var tokensWithError = tokens
                .Where(t => t.StartIndex < errorStartIndex)
                .Append(new Token(Token.ErrorSymbolId, errorStartIndex, 0)).ToList();
            return tokensWithError;
        }

        tokens.Add(Token.EofSymbol(formula.Length));
        return tokens;
    }

    private class LexerErrorListener : IAntlrErrorListener<int>
    {
        internal int? ErrorStartIndex { get; private set; }

        public void SyntaxError(TextWriter output, IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            // Params don't provide access to the stream char index property directly, so pass it through 
            ErrorStartIndex ??= ((Lexer)recognizer).TokenStartCharIndex;
        }
    }
}