using ClosedXML.Parser.Pratt;

namespace ClosedXML.Parser.Tests.Lexers;

public class LexerTests
{
    // ( [0-9] )+
    [InlineData("0")]
    [InlineData("1")]
    [InlineData("90")]
    [InlineData("00050")]

    // ( [0-9] )+ '.' ( [0-9]] )+
    [InlineData("0.0")]
    [InlineData("1.2")]
    [InlineData("0010.0020")]
    [InlineData("999.99")]

    // '.' ( [0-9]] )+
    [InlineData(".0")]
    [InlineData(".1")]
    [InlineData(".0001")]
    [InlineData(".987")]

    // ( [0-9] )+ [Ee] ( [0-9]] )+
    [InlineData("0e0")]
    [InlineData("0E0")]
    [InlineData("1e2")]
    [InlineData("1E2")]
    [InlineData("987e12")]

    // ( [0-9] )+ '.' ( [0-9]] )+ [Ee] ( [0-9]] )+
    [InlineData("0.0e4")]
    [InlineData("12.724e13")]
    [InlineData("12.3E2")]

    // '.' ( [0-9]] )+ [Ee] ( [0-9]] )+
    [InlineData(".0e0")]
    [InlineData(".1e2")]
    [InlineData(".987e54")]

    // ( [0-9] )+ [Ee] [+-] ( [0-9]] )+
    [InlineData("1e+7")]
    [InlineData("74e-32")]
    [InlineData("15E-0")]
    [InlineData("0e+0")]
    [InlineData("01e+7")]

    // ( [0-9] )+ '.' ( [0-9]] )+ [Ee] [+-] ( [0-9]] )+
    [InlineData("0.0e+0")]
    [InlineData("1.2e+3")]
    [InlineData("01.2e+3")]
    [InlineData("1.2E+3")]
    [InlineData("12.34e+56")]

    // '.' ( [0-9]] )+ [Ee] [+-] ( [0-9]] )+
    [InlineData(".0e+0")]
    [InlineData(".1e+2")]
    [InlineData(".12E+34")]
    [InlineData(".012e+034")]
    [Theory]
    public void Number_ok(string input)
    {
        AssertToken(TokenType.Number, input);
    }

    [InlineData("0e+")]
    [InlineData(".0e+")]
    [Theory]
    public void Number_fails(string input)
    {
        AssertFail(input, "Number");
    }

    [InlineData("\"\"")]
    [InlineData("\"Some text\"")]
    [InlineData("\"Some \"\" text\"")]
    [InlineData("\"\uD83E\uDD8A\"")] // Fox face through surrogates
    [Theory]
    public void Text_ok(string input)
    {
        AssertToken(TokenType.Text, input);
    }

    [InlineData("\"")]
    [InlineData("\"text")]
    [InlineData("\"text\"\"")]
    [InlineData("\"Some \"\" text")]
    [Theory]
    public void Text_must_be_terminated(string input)
    {
        AssertFail(input, "unterminated literal");
    }

    [InlineData("\"\u0015\"")]
    [Theory]
    public void Text_must_be_contain_xml_10_characters(string input)
    {
        AssertFail(input, "Invalid text character");
    }

    [InlineData("#DIV/0!")]
    [InlineData("#GETTING_DATA")]
    [InlineData("#N/A")]
    [InlineData("#NAME?")]
    [InlineData("#NULL!")]
    [InlineData("#NUM!")]
    [InlineData("#REF!")]
    [InlineData("#VALUE!")]
    [InlineData("#ref!")]
    [Theory]
    public void Error_ok(string input)
    {
        AssertToken(TokenType.Error, input);
    }

    [Fact]
    public void Lexer_throws_on_unpaired_surrogates()
    {
        // Either Visual Studio or NUnit is converting invalid surrogates to -1/65536. O
        var invalidCodeUnits = new[]
        {
            "\uD83E", // Unpaired high surrogate for Fox Face
            "\uD83E*", // Unpaired high surrogate for Fox Face
            "\uDD8A", // Unpaired low surrogate for Fox Face
            "\uDD8A*\"", // Unpaired low surrogate for Fox Face
            "\uDD8A\uD83E", // Low surrogate first
        };
        foreach (var invalidText in invalidCodeUnits)
        {
            AssertFail(invalidText, "surrogate");
        }
    }

    [InlineData("''")]
    [InlineData("'[1]Something'")]
    [InlineData("'Jane''s'")]
    [InlineData("'New York'")]
    [InlineData("'January 1st:December 31st'")]
    [InlineData("'[7]Year 20:Year 25'")]
    [InlineData("'[Book.xlsx]Year 20:Year 25'")]
    [InlineData("'[End*Near.xlsx]Final'")]
    [InlineData("''''''")]
    [Theory]
    public void QIdent_ok(string input)
    {
        AssertToken(TokenType.QIdent, input);
    }

    [InlineData("'")]
    [InlineData("'Jane''s")]
    [InlineData("'''''")]
    [Theory]
    public void QIdent_must_be_terminated(string input)
    {
        AssertFail(input, "unterminated literal");
    }

    [InlineData("ABC")]
    [InlineData("A1")]
    [InlineData("$A$1")]
    [InlineData("AEF$A$1")]
    [InlineData("name")]
    [InlineData("TRUE")]
    [InlineData("FALSE")]
    [InlineData("true")]
    [InlineData("false")]
    [InlineData("?name")]
    [InlineData("\\name")]
    [InlineData("_name")]
    [InlineData("name?")]
    [InlineData("name\\")]
    [InlineData("name_")]
    [InlineData("some.name")]
    [InlineData("_xlfn.ACOT")]
    [InlineData("\u05D0\u05D1\u05E0")] // stone in hebrew - Letters from other languages
    [InlineData("\u05E9\u05B0\u05DC\u05D5\u05DD")] // shalom - A mark from other languages
    [Theory]
    public void Ident_ok(string input)
    {
        AssertToken(TokenType.Ident, input);
    }

    [Fact]
    public void Ident_stops_at_operators()
    {
        var operators = new Dictionary<TokenType, string>
        {
            { TokenType.Bang, "!" },
            { TokenType.Comma, "," },
            { TokenType.Semicolon, ";" },
            { TokenType.Pow, "^" },
            { TokenType.Mul, "*" },
            { TokenType.Div, "/" },
            { TokenType.Plus, "+" },
            { TokenType.Minus, "-" },
            { TokenType.Concat, "&" },
            { TokenType.Equal, "=" },
            { TokenType.NotEqual, "<>" },
            { TokenType.Less, "<" },
            { TokenType.LessEqual, "<=" },
            { TokenType.Greater, ">" },
            { TokenType.GreaterEqual, ">=" },
            { TokenType.Percent, "%" },
            { TokenType.Range, ":" },
            { TokenType.Spill, "#" },
            { TokenType.Intersection, "@" },
            { TokenType.LeftParen, "(" },
            { TokenType.RightParen, ")" },
            { TokenType.LeftCurly, "{" },
            { TokenType.RightCurly, "}" },
            { TokenType.Whitespace, " " },
        };

        foreach (var (opType, opText) in operators)
        {
            var input = "name" + opText;
            var lexer = new Lexer(input);

            var identToken = lexer.Consume();
            Assert.Equal(TokenType.Ident, identToken.Type);
            Assert.Equal("name", identToken.GetText(input).ToString());

            var opToken = lexer.Consume();
            Assert.Equal(opType, opToken.Type);
            Assert.Equal(opText, opToken.GetText(input).ToString());
        }
    }

    [InlineData("[1]")]
    [InlineData("[]")]
    [InlineData("['[]")]
    [InlineData("[Book1.xlsx]")]
    [InlineData("[#Data]")]
    [InlineData("[[#Data]]")]
    [InlineData("[[#Data],[#Headers]]")]
    [InlineData("['#]")]
    [InlineData("[985]")]
    [InlineData("[Jan:Dec]")]
    [InlineData("['['['[]")]
    [InlineData("[']']']]")]
    [Theory]
    public void SquareIdent_ok(string input)
    {
        AssertToken(TokenType.SquareIdent, input);
    }

    [InlineData("[Ja[[a]]]")]
    [Theory]
    public void SquareIdent_at_most_two_nested_brackets(string input)
    {
        // Mostly to keep within something DFA can do.
        AssertFail(input, "at most two nested square brackets");
    }

    [InlineData("[")]
    [InlineData("[text")]
    [InlineData("[[")]
    [InlineData("[a[b")]
    [InlineData("[Start[]and end")]
    [InlineData("[Start[']and end']")]
    [Theory]
    public void SquareIdent_must_be_paired(string input)
    {
        AssertFail(input, "Unable to find closing square bracket");
    }

    [InlineData((int)TokenType.Bang, "!")]
    [InlineData((int)TokenType.Range, ":")]
    [InlineData((int)TokenType.Comma, ",")]
    [InlineData((int)TokenType.Semicolon, ";")]
    [InlineData((int)TokenType.Pow, "^")]
    [InlineData((int)TokenType.Mul, "*")]
    [InlineData((int)TokenType.Div, "/")]
    [InlineData((int)TokenType.Plus, "+")]
    [InlineData((int)TokenType.Minus, "-")]
    [InlineData((int)TokenType.Concat, "&")]
    [InlineData((int)TokenType.Percent, "%")]
    [InlineData((int)TokenType.Spill, "#")]
    [InlineData((int)TokenType.Intersection, "@")]
    [InlineData((int)TokenType.LeftParen, "(")]
    [InlineData((int)TokenType.RightParen, ")")]
    [InlineData((int)TokenType.LeftCurly, "{")]
    [InlineData((int)TokenType.RightCurly, "}")]
    [Theory]
    public void Single_char_tokens_ok(int token, string input)
    {
        // TODO: Dump xUnit. Can't even use internal classes as test fixtures, so I have to pass enum as int.
        AssertToken((TokenType)token, input);
    }

    [InlineData((int)TokenType.Equal, "=")]
    [InlineData((int)TokenType.NotEqual, "<>")]
    [InlineData((int)TokenType.Less, "<")]
    [InlineData((int)TokenType.LessEqual, "<=")]
    [InlineData((int)TokenType.Greater, ">")]
    [InlineData((int)TokenType.GreaterEqual, ">=")]
    [Theory]
    public void Comparison_tokens_ok(int token, string input)
    {
        AssertToken((TokenType)token, input);
    }

    [InlineData("\t")]
    [InlineData("\n")]
    [InlineData("\r")]
    [InlineData(" ")]
    [InlineData("\t \r\n")]
    [Theory]
    public void Whitespace_ok(string input)
    {
        AssertToken(TokenType.Whitespace, input);
    }

    private static void AssertToken(TokenType type, string input)
    {
        var lexer = new Lexer(input);
        var token = lexer.Consume();
        Assert.Equal(type, token.Type);
        Assert.Equal(input, token.GetText(input).ToString());
    }

    private static void AssertFail(string input, string exceptionSubstring)
    {
        var lexer = new Lexer(input);
        var exception = Assert.Throws<ParsingException>(() => lexer.Consume());
        Assert.NotNull(exception);
        Assert.True(exception.Message.Contains(exceptionSubstring), $"Expected to find '{exceptionSubstring}', but not found in '{exception.Message}'.");
    }
}
