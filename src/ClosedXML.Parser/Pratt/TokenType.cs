namespace ClosedXML.Parser.Pratt;

/// <summary>
/// A token types for a lexer.
/// </summary>
internal enum TokenType
{
    /// <summary>
    /// An identifier in a formula. In most generic form, it's a name.
    /// Essentially a text that doesn't start with a number without a whitespace.
    /// Includes following rules from ABNF:
    /// <list type="bullet">
    ///   <item>A1-column, A1-row, A1-cell</item>
    ///   <item>name</item>
    ///   <item>logical-constant</item>
    ///   <item>sheet-name</item>
    /// </list>
    /// </summary>
    /// <remarks>
    /// <para>
    /// Excel doesn't have clear distinction between identifiers and keywords or other things.
    /// Example: <c>LOG10</c> could be an A1 reference name (column <c>LOG</c>, row <c>10</c>),
    /// a function (<c>LOG10(14)</c>) or sheet name (<c>LOG10!A1</c>). To determine it, parser
    /// needs a context. And context isn't available in the lexer.
    /// </para>
    /// <para>
    /// Unlike A1, R1C1 has to be recognized in a parser. The minus sign along with square
    /// brackets in a <c>R[-1]C[-1]</c> is a deal-breaker.
    /// </para>
    /// </remarks>
    Ident, // TODO: sheet-name rule description is obviously not true, needs to be checked manually

    /// <summary>
    /// A floating point number. The textual representation isn't limited number to maximum
    /// precision of IEEE 754 standard.
    /// </summary>
    Number,

    /// <summary>
    /// A text inside double quotes. Double quotes are escaped by doubling.
    /// </summary>
    Text,

    /// <summary>
    /// Error literal, e.g. <c>#N/A</c>.
    /// </summary>
    Error,

    /// <summary>
    /// A token representing text between two single quotes. Single quotes are escaped by doubling.
    /// <list type="bullet">
    ///  <item><c>'Jane''s'</c> - sheet names with escaped character</item>
    ///  <item><c>'New York'</c> - sheet names with spaces</item>
    ///  <item><c>'January 1st:December 31st'</c> - 3D references of sheets with spaces</item>
    ///  <item><c>'[7]Year 20:Year 25'</c> - 3D references to external workbook.</item>
    ///  <item><c>'[Book.xlsx]Year 20:Year 25'</c> - 3D references to external workbook.</item>
    /// </list>
    /// </summary>
    /// <remarks>
    /// The ABNF says 
    /// <code>
    ///   sheet-name-special = sheet-name-base-character [*sheet-name-character-special sheet-name-base-character]
    ///   sheet-name-character-special = 2apostrophe / sheet-name-base-character
    ///   sheet-name-base-character = character; MUST NOT be ', *, [, ], \, :, /, ?, or Unicode character 'END OF TEXT'
    ///   character = as defined by the production Char in the [W3C-XML] section 2.2
    /// </code>
    /// but we accept everything in lexer. The <c>[</c> and <c>]</c> must be part of it due
    /// to workbook index or <c>*</c> could be a valid name of a workbook. Since is has to be
    /// filtered in the parser anyway, don't burden the lexer.
    /// </remarks>
    QIdent,

    /// <summary>
    /// A span of content inside a square brackets. The token inside brackets includes escaped
    /// brackets and structure reference keywords.
    /// </summary>
    /// <remarks>
    /// <para>
    /// There is a problem with it being either book
    /// reference or structure reference. In addition, we might want to parse names of book files
    /// in the future. Lexer should be doable through DFA and this is really hard, so just detect
    /// token and leave decision to the parser. Nested square brackets are not allowed (must be
    /// escaped), so there are at most two level deep nested brackets (<c>[[#Header],[#Data]]</c>),
    /// which is doable by DFA (unlimited nesting isn't).
    /// </para>
    /// <para>
    /// Examples:
    /// <list type="bullet">
    ///  <item><c>[1]</c> - either structure reference to a column '1' or book index.</item>
    ///  <item><c>[]</c> - structure reference to a whole table (from first to last).</item>
    ///  <item><c>['[]</c> - structure reference to a column '['.</item>
    ///  <item><c>[Book1.xlsx]</c> - book reference, not part of official grammar.</item>
    ///  <item><c>[#Data]</c> - structure reference to data portion of a table</item>
    ///  <item><c>[[#Data]]</c> - Nested reference</item>
    ///  <item><c>['#]</c> - structure reference to a column '#'</item>
    /// </list>
    /// </para>
    /// </remarks>
    SquareIdent,

    /// <summary>
    /// Bang <c>!</c>. It is used in sheet reference, bang names and bang references.
    /// </summary>
    Bang,

    // Operators
    /// <summary>
    /// <c>,</c> - argument separator in function call, range union operator, or separator of
    /// values in a row for an array literal.
    /// </summary>
    Comma,

    /// <summary>
    /// <c>;</c> - separator of rows in array literal.
    /// </summary>
    Semicolon,

    /// <summary>
    /// <c>^</c> - power operator.
    /// </summary>
    Pow,

    /// <summary>
    /// <c>*</c> - multiplication operator.
    /// </summary>
    Mul,

    /// <summary>
    /// <c>/</c> - division operator.
    /// </summary>
    Div,

    /// <summary>
    /// <c>-</c> - prefix or binary plus operator.
    /// </summary>
    Plus,

    /// <summary>
    /// <c>-</c> - prefix or binary minus operator.
    /// </summary>
    Minus,

    /// <summary>
    /// <c>&amp;</c> - text concatenation operator.
    /// </summary>
    Concat,

    /// <summary>
    /// <c>=</c> equal comparison operator.
    /// </summary>
    Equal,

    /// <summary>
    /// <c>&lt;&gt;</c> not equals comparison operator.
    /// </summary>
    NotEqual,

    /// <summary>
    /// <c>&lt;</c> less than comparison operator.
    /// </summary>
    Less,

    /// <summary>
    /// <c>&lt;=</c> less than or equal comparison operator.
    /// </summary>
    LessEqual,

    /// <summary>
    /// <c>&gt;</c> greater than comparison operator.
    /// </summary>
    Greater,

    /// <summary>
    /// <c>&gt;=</c> greater than or equal comparison operator.
    /// </summary>
    GreaterEqual,

    /// <summary>
    /// <c>%</c> postfix operator.
    /// </summary>
    Percent,

    /// <summary>
    /// <c>:</c> - range of two references.
    /// </summary>
    Range,

    /// <summary>
    /// <c>#</c> - postfix reference operator.
    /// </summary>
    Spill,

    /// <summary>
    /// <c>@</c> - implicit intersection of reference.
    /// </summary>
    Intersection,

    /// <summary>
    /// <c>(</c> - a nested group operator or opening parenthesis of a function call.
    /// </summary>
    LeftParen,

    /// <summary>
    /// <c>)</c> - a nested group operator or closing parenthesis of a function call.
    /// </summary>
    RightParen,

    /// <summary>
    /// <c>{</c> - opening token of array literal.
    /// </summary>
    LeftCurly,

    /// <summary>
    /// <c>}</c> - closing token of array literal.
    /// </summary>
    RightCurly,

    /// <summary>
    /// <c> </c> - binary intersection operator or whitespace that will be ignored by parser.
    /// </summary>
    Whitespace,

    /// <summary>
    /// End of file.
    /// </summary>
    Eof,
}
