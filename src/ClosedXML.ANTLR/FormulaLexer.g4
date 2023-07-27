lexer grammar FormulaLexer;

/* --------------------------------- Errors -------------------------------- */
REF_CONSTANT
        : '#' R E F '!'
        ;

NONREF_ERRORS
        : '#' D I V '/0!'
        | '#' N '/' A
        | '#' N A M E '?'
        | '#' N U L L '!'
        | '#' N U M '!'
        | '#' V A L U E '!'
        | '#' G E T T I N G '_' D A T A
        ;

/* ------------------------- Logical constant ------------------------------ */

LOGICAL_CONSTANT
        : F A L S E
        | T R U E
        ;

/* -------------------------- Number constant ------------------------------ */

/*
 * Token doesn't have '-', because it would eat the unary/binary op, e.g. 1-2,
 * where there would be a two NUMERICAL_CONSTANT tokens without operation.
 */
NUMERICAL_CONSTANT
        : SIGNIFICAND_PART EXPONENT_PART?
        ;

fragment SIGNIFICAND_PART
        : WHOLE_NUMBER_PART (FRACTIONAL_PART)?
        | FRACTIONAL_PART
        ;

fragment WHOLE_NUMBER_PART
        : DIGIT_SEQUENCE
        ;

fragment FRACTIONAL_PART
        : FULL_STOP DIGIT_SEQUENCE
        ;

fragment EXPONENT_PART
        : EXPONENT_CHARACTER ( SIGN )? DIGIT_SEQUENCE
        ;

fragment FULL_STOP
        : '.'
        ;

fragment SIGN
        : '+' | '-'
        ;
fragment EXPONENT_CHARACTER
        : E
        ;

fragment DIGIT_SEQUENCE
        : ( DECIMAL_DIGIT )+
        ;

fragment DECIMAL_DIGIT
        : [0-9]
        ;

fragment NONZERO_DECIMAL_DIGIT
        : [1-9]
        ;

/* -------------------------- String constant ------------------------------ */

STRING_CONSTANT
        : DOUBLE_QUOTE STRING_CHARS? DOUBLE_QUOTE
        ;

fragment STRING_CHARS
        : STRING_CHAR (STRING_CHAR)*
        ;

 // MUST NOT be a double-quote
fragment STRING_CHAR
        : ( ESCAPED_DOUBLE_QUOTE | CHARCATER_WITHOUT_DOUBLE_QUOTE )
        ;
fragment ESCAPED_DOUBLE_QUOTE
        :  '""'
        ;

fragment DOUBLE_QUOTE
        : '"'
        ;
// Character as defined by the production Char in the [W3C-XML] section 2.2
fragment CHARACTER
        : '\u0009'
        | '\u000A'
        | '\u000D'
        | '\u0020' .. '\uD7FF'
        | '\uE000' .. '\uFFFD'
        | '\u{10000}' .. '\u{10FFFF}'
        ;

fragment CHARCATER_WITHOUT_DOUBLE_QUOTE
        : '\u0009'
        | '\u000A'
        | '\u000D'
        | '\u0020' .. '\u0021'
        | '\u0023' .. '\uD7FF'
        | '\uE000' .. '\uFFFD'
        | '\u{10000}' .. '\u{10FFFF}'
        ;

/* ----------------------------- Operators --------------------------------- */

POW: WHITESPACES '^' WHITESPACES;
MULT: WHITESPACES '*' WHITESPACES;
DIV: WHITESPACES '/' WHITESPACES;
PLUS: WHITESPACES '+' WHITESPACES;
MINUS: WHITESPACES '-' WHITESPACES;
CONCAT: WHITESPACES '&' WHITESPACES;
EQUAL: WHITESPACES '=' WHITESPACES;
NOT_EQUAL: WHITESPACES '<>' WHITESPACES;
LESS_OR_EQUAL_THAN: WHITESPACES '<=' WHITESPACES;
LESS_THAN: WHITESPACES '<' WHITESPACES;
GREATER_OR_EQUAL_THAN: WHITESPACES '>=' WHITESPACES;
GREATER_THAN: WHITESPACES '>' WHITESPACES;
PERCENT: WHITESPACES '%' WHITESPACES;
SEMICOLON: WHITESPACES ';' WHITESPACES;
COLON: WHITESPACES ':' WHITESPACES;
OPEN_BRACE: WHITESPACES '(' WHITESPACES;
CLOSE_BRACE: WHITESPACES ')' WHITESPACES;
OPEN_CURLY: WHITESPACES '{' WHITESPACES;
CLOSE_CURLY: WHITESPACES '}' WHITESPACES;
COMMA : WHITESPACES ',' WHITESPACES;
SPACE: WHITESPACES ' ' WHITESPACES;
OPEN_SQUARE: WHITESPACES '[' WHITESPACES;
CLOSE_SQUARE: WHITESPACES ']' WHITESPACES;

fragment WHITESPACES : (' ' | '\u000D' | '\u000A')*;

/* -------------------------- External References -------------------------- */

BOOK_PREFIX : WORKBOOK_INDEX '!';

// MS-XLSX 2.2.2.1: The formula MUST NOT use the bang-reference or bang-name.
BANG_REFERENCE
        : '!' (A1_REFERENCE | REF_CONSTANT)
        ;

SHEET_RANGE_PREFIX
        : SHEET_RANGE '!'
        ;

SINGLE_SHEET_PREFIX
        : SINGLE_SHEET '!'
        ;

// This rule is not used for cell formulas, but is used for sparklines (possibly others)
//single-sheet-area = single-sheet-prefix A1-area
fragment SINGLE_SHEET
                : (WORKBOOK_INDEX)? SHEET_NAME
                // because of ambiguity of SHEET_NAME and SHEET_NAME_SPECIAL, add alternative
                | TICK (WORKBOOK_INDEX)? (SHEET_NAME_SPECIAL | SHEET_NAME) TICK
                ;

fragment SHEET_RANGE
                : WORKBOOK_INDEX? SHEET_NAME ':' SHEET_NAME
                | TICK WORKBOOK_INDEX? SHEET_NAME_SPECIAL ':' SHEET_NAME_SPECIAL TICK
                ;

fragment WORKBOOK_INDEX
        : '[' WHOLE_NUMBER_PART ']'
        ;

fragment SHEET_NAME
        : SHEET_NAME_CHARACTERS
        ;

fragment SHEET_NAME_CHARACTERS
        : SHEET_NAME_CHARACTER+
        ;

// The original didn't specify that  various chars, like '!', '(', ')', should
// not be part of sheet name => chars were removed from the fragment.
// Original comment: Must not be operator, ', [, ], \, or ?
fragment SHEET_NAME_CHARACTER
        : '\u0009'
        | '\u000A'
        | '\u000D'
        | '\u0022' .. '\u0024' // 0020 [space] 0021 ! 0025 % 0026 & 0027 ' 0028 ( 0029 ) 002A * 002B + 002C , 002D -
        | '\u002E'             // 002F /
        | '\u0030' .. '\u0039' // 003A :
        | '\u003B'             // 003C < 003D = 003E > 003F ?
        | '\u0040' .. '\u005A' // 005B [ 005C \ 005D ] 005E ^
        | '\u005F' .. '\uD7FF'
        | '\uE000' .. '\uFFFD'
        | '\u{10000}' .. '\u{10FFFF}'
        ;

fragment SHEET_NAME_SPECIAL
        : SHEET_NAME_BASE_CHARACTER+ ((SHEET_NAME_CHARACTER_SPECIAL)* SHEET_NAME_BASE_CHARACTER)?
        ;

fragment SHEET_NAME_CHARACTER_SPECIAL
        : TICK TICK
        | SHEET_NAME_BASE_CHARACTER
        ;

// MUST NOT be ', *, [, ], \, :, /, ?, or Unicode character 'END OF TEXT'
fragment SHEET_NAME_BASE_CHARACTER
        : '\u0009'
        | '\u000A'
        | '\u000D'
        | '\u0020' .. '\u0026' // 0027 '
        | '\u0028' .. '\u0029' // 002A *
        | '\u002B' .. '\u002E' // 002F /
        | '\u0030' .. '\u0039' // 003A :
        | '\u003B' .. '\u003E' // 003F ?
        | '\u0040' .. '\u005A' // 005B [ 005C \ 005D ]
        | '\u005E' .. '\uD7FF'
        | '\uE000' .. '\uFFFD'
        | '\u{10000}' .. '\u{10FFFF}'
        ;

/* -------------------------- Local A1 References -------------------------- */

A1_REFERENCE
        : A1_COLUMN ':' A1_COLUMN
        | A1_ROW ':' A1_ROW
        | A1_CELL
        | A1_AREA
        ;

fragment A1_CELL
        : A1_COLUMN A1_ROW
        ;

fragment A1_AREA
        : A1_CELL ':' A1_CELL
        ;

fragment A1_COLUMN
        : A1_RELATIVE_COLUMN | A1_ABSOLUTE_COLUMN
        ;

fragment A1_RELATIVE_COLUMN
        : COLUMN_LETTER
        | COLUMN_LETTER COLUMN_LETTER
        | A_to_W COLUMN_LETTER COLUMN_LETTER
        | X A_to_E COLUMN_LETTER
        | X F A_to_D
        ;

fragment A_to_D
        : [A-Da-d]
        ;

fragment A_to_E
        : [A-Ea-e]
        ;

fragment A_to_W
        : [A-Wa-w]
        ;

fragment COLUMN_LETTER
        : [A-Za-z]
        ;

fragment A1_ABSOLUTE_COLUMN
        : '$' A1_RELATIVE_COLUMN
        ;

fragment A1_ROW
        : A1_RELATIVE_ROW
        | A1_ABSOLUTE_ROW
        ;

fragment A1_RELATIVE_ROW
        : ROW_DIGIT_SEQUENCE
        ;

fragment ROW_DIGIT_SEQUENCE
        : NONZERO_DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '10' [0-3] DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '104' [0-7] DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '1048' [0-4] DECIMAL_DIGIT DECIMAL_DIGIT
        | '10485' [0-6] DECIMAL_DIGIT
        | '104857' [0-6]
        ;

fragment A1_ABSOLUTE_ROW
        : '$' A1_RELATIVE_ROW
        ;

/* -------------------------- Local R1C1 References -------------------------- */
/*
A1_REFERENCE
        : R1C1_COLUMN ':' R1C1_COLUMN
        | R1C1_ROW ':' R1C1_ROW
        | A1_CELL
        | R1C1_AREA
        ;

fragment A1_CELL
        : R1C1_ROW R1C1_COLUMN
        ;

fragment R1C1_AREA
        : A1_CELL ':' A1_CELL
        ;

fragment R1C1_COLUMN
        : R1C1_RELATIVE_COLUMN | R1C1_ABSOLUTE_COLUMN
        ;

fragment R1C1_RELATIVE_COLUMN
        : 'C[' ('-')? COLUMN_RELATIVE_DIGIT_SEQUENCE ']'
        ;

fragment R1C1_ABSOLUTE_COLUMN
        : 'C' COLUMN_DIGIT_SEQUENCE
        ;

fragment R1C1_ROW
        : R1C1_RELATIVE_ROW
        | R1C1_ABSOLUTE_ROW
        ;

fragment R1C1_RELATIVE_ROW
        : 'R[' ('-')? ROW_RELATIVE_DIGIT_SEQUENCE ']'
        ;

fragment R1C1_ABSOLUTE_ROW
        : 'R' ROW_DIGIT_SEQUENCE
        ;

fragment ROW_DIGIT_SEQUENCE
        : ROW_RELATIVE_DIGIT_SEQUENCE
        | '1048576'
        ;

fragment ROW_RELATIVE_DIGIT_SEQUENCE
        : NONZERO_DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '10' [0-3] DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '104' [0-7] DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '1048' [0-4] DECIMAL_DIGIT DECIMAL_DIGIT
        | '10485' [0-6] DECIMAL_DIGIT
        | '104857' [0-5]
        ;

fragment COLUMN_DIGIT_SEQUENCE
        : COLUMN_RELATIVE_DIGIT_SEQUENCE
        | '16384'
        ;

fragment COLUMN_RELATIVE_DIGIT_SEQUENCE
        : NONZERO_DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | NONZERO_DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '1' [0-5]  DECIMAL_DIGIT DECIMAL_DIGIT DECIMAL_DIGIT
        | '16' [0-2]  DECIMAL_DIGIT DECIMAL_DIGIT
        | '163' [0-7]  DECIMAL_DIGIT
        | '1638' [0-3]
        ;

/* ------------------------------ Functions -------------------------------- */

// Ref must be before function to keep priority
REF_FUNCTION_LIST
        : ( C H O O S E | I F | I N D E X | I N D I R E C T | O F F S E T ) '(' WHITESPACES
        ;

/*
 * Comment out built in functions to make lexer smaller. All built-in names can
 * be detected by NAME token and we have to match function name to a function
 * reference anyway. Also, most users enter future function without prefix and
 * we have to match it correctly.
 */
/*
FUNCTION_LIST
        : ('ABS' | 'ABSREF' | 'ACCRINT' | 'ACCRINTM' | 'ACOS' | 'ACOSH' | 'ACTIVE.CELL' | 'ADD.BAR' | 'ADD.COMMAND' | 'ADD.MENU' | 'ADD.TOOLBAR' | 'ADDRESS' | 'AMORDEGRC' | 'AMORLINC' | 'AND' | 'APP.TITLE' | 'AREAS' | 'ARGUMENT' | 'ASC' | 'ASIN' | 'ASINH' | 'ATAN' | 'ATAN2' | 'ATANH' | 'AVEDEV' | 'AVERAGE' | 'AVERAGEA' | 'AVERAGEIF' | 'AVERAGEIFS' | 'BAHTTEXT' | 'BESSELI' | 'BESSELJ' | 'BESSELK' | 'BESSELY' | 'BETADIST' | 'BETAINV' | 'BIN2DEC' | 'BIN2HEX' | 'BIN2OCT' | 'BINOMDIST' | 'BREAK' | 'CALL' | 'CALLER' | 'CANCEL.KEY' | 'CEILING' | 'CELL' | 'CHAR' | 'CHECK.COMMAND' | 'CHIDIST' | 'CHIINV' | 'CHITEST' | 'CLEAN' | 'CODE' | 'COLUMN' | 'COLUMNS' | 'COMBIN' | 'COMPLEX' | 'CONCAT' | 'CONCATENATE' | 'CONFIDENCE' | 'CONVERT' | 'CORREL' | 'COS' | 'COSH' | 'COUNT' | 'COUNTA' | 'COUNTBLANK' | 'COUNTIF' | 'COUNTIFS' | 'COUPDAYBS' | 'COUPDAYS' | 'COUPDAYSNC' | 'COUPNCD' | 'COUPNUM' | 'COUPPCD' | 'COVAR' | 'CREATE.OBJECT' | 'CRITBINOM' | 'CUBEKPIMEMBER' | 'CUBEMEMBER' | 'CUBEMEMBERPROPERTY' | 'CUBERANKEDMEMBER' | 'CUBESET' | 'CUBESETCOUNT' | 'CUBEVALUE' | 'CUMIPMT' | 'CUMPRINC' | 'CUSTOM.REPEAT' | 'CUSTOM.UNDO' | 'DATE' | 'DATEDIF' | 'DATESTRING' | 'DATEVALUE' | 'DAVERAGE' | 'DAY' | 'DAYS360' | 'DB' | 'DBCS' | 'DCOUNT' | 'DCOUNTA' | 'DDB' | 'DEC2BIN' | 'DEC2HEX' | 'DEC2OCT' | 'DEGREES' | 'DELETE.BAR' | 'DELETE.COMMAND' | 'DELETE.MENU' | 'DELETE.TOOLBAR' | 'DELTA' | 'DEREF' | 'DEVSQ' | 'DGET' | 'DIALOG.BOX' | 'DIRECTORY' | 'DISC' | 'DMAX' | 'DMIN' | 'DOCUMENTS' | 'DOLLAR' | 'DOLLARDE' | 'DOLLARFR' | 'DPRODUCT' | 'DSTDEV' | 'DSTDEVP' | 'DSUM' | 'DURATION' | 'DVAR' | 'DVARP' | 'ECHO' | 'EDATE' | 'EFFECT' | 'ELSE' | 'ELSE.IF' | 'ENABLE.COMMAND' | 'ENABLE.TOOL' | 'END.IF' | 'EOMONTH' | 'ERF' | 'ERFC' | 'ERROR' | 'ERROR.TYPE' | 'EVALUATE' | 'EVEN' | 'EXACT' | 'EXEC' | 'EXECUTE' | 'EXP' | 'EXPONDIST' | 'FACT' | 'FACTDOUBLE' | 'FALSE' | 'FCLOSE' | 'FDIST' | 'FILES' | 'FIND' | 'FINDB' | 'FINV' | 'FISHER' | 'FISHERINV' | 'FIXED' | 'FLOOR' | 'FOPEN' | 'FOR' | 'FOR.CELL' | 'FORECAST' | 'FORMULA.CONVERT' | 'FPOS' | 'FREAD' | 'FREADLN' | 'FREQUENCY' | 'FSIZE' | 'FTEST' | 'FV' | 'FVSCHEDULE' | 'FWRITE' | 'FWRITELN' | 'GAMMADIST' | 'GAMMAINV' | 'GAMMALN' | 'GCD' | 'GEOMEAN' | 'GESTEP' | 'GET.BAR' | 'GET.CELL' | 'GET.CHART.ITEM' | 'GET.DEF' | 'GET.DOCUMENT' | 'GET.FIELD' | 'GET.FORMULA' | 'GET.ITEM' | 'GET.LINK.INFO' | 'GET.MOVIE' | 'GET.NAME' | 'GET.NOTE' | 'GET.OBJECT' | 'GET.TOOL' | 'GET.TOOLBAR' | 'GET.VIEW' | 'GET.WINDOW' | 'GET.WORKBOOK' | 'GET.WORKSPACE' | 'GETPIVOTDATA' | 'GOTO' | 'GROUP' | 'GROWTH' | 'HALT' | 'HARMEAN' | 'HELP' | 'HEX2BIN' | 'HEX2DEC' | 'HEX2OCT' | 'HLOOKUP' | 'HOUR' | 'HYPERLINK' | 'HYPGEOMDIST' | 'IFS' | 'IFERROR' | 'IMABS' | 'IMAGINARY' | 'IMARGUMENT' | 'IMCONJUGATE' | 'IMCOS' | 'IMDIV' | 'IMEXP' | 'IMLN' | 'IMLOG10' | 'IMLOG2' | 'IMPOWER' | 'IMPRODUCT' | 'IMREAL' | 'IMSIN' | 'IMSQRT' | 'IMSUB' | 'IMSUM' | 'INFO' | 'INITIATE' | 'INPUT' | 'INT' | 'INTERCEPT' | 'INTRATE' | 'IPMT' | 'IRR' | 'ISBLANK' | 'ISERR' | 'ISERROR' | 'ISEVEN' | 'ISLOGICAL' | 'ISNA' | 'ISNONTEXT' | 'ISNUMBER' | 'ISODD' | 'ISPMT' | 'ISREF' | 'ISTEXT' | 'ISTHAIDIGIT' | 'KURT' | 'LARGE' | 'LAST.ERROR' | 'LCM' | 'LEFT' | 'LEFTB' | 'LEN' | 'LENB' | 'LINEST' | 'LINKS' | 'LN' | 'LOG' | 'LOG10' | 'LOGEST' | 'LOGINV' | 'LOGNORMDIST' | 'LOOKUP' | 'LOWER' | 'MATCH' | 'MAX' | 'MAXA' | 'MAXIFS' | 'MDETERM' | 'MDURATION' | 'MEDIAN' | 'MID' | 'MIDB' | 'MIN' | 'MINA' | 'MINIFS' | 'MINUTE' | 'MINVERSE' | 'MIRR' | 'MMULT' | 'MOD' | 'MODE' | 'MONTH' | 'MOVIE.COMMAND' | 'MROUND' | 'MULTINOMIAL' | 'N' | 'NA' | 'NAMES' | 'NEGBINOMDIST' | 'NETWORKDAYS' | 'NEXT' | 'NOMINAL' | 'NORMDIST' | 'NORMINV' | 'NORMSDIST' | 'NORMSINV' | 'NOT' | 'NOTE' | 'NOW' | 'NPER' | 'NPV' | 'NUMBERSTRING' | 'OCT2BIN' | 'OCT2DEC' | 'OCT2HEX' | 'ODD' | 'ODDFPRICE' | 'ODDFYIELD' | 'ODDLPRICE' | 'ODDLYIELD' | 'OPEN.DIALOG' | 'OPTIONS.LISTS.GET' | 'OR' | 'PAUSE' | 'PEARSON' | 'PERCENTILE' | 'PERCENTRANK' | 'PERMUT' | 'PHONETIC' | 'PI' | 'PMT' | 'POISSON' | 'POKE' | 'POWER' | 'PPMT' | 'PRESS.TOOL' | 'PRICE' | 'PRICEDISC' | 'PRICEMAT' | 'PROB' | 'PRODUCT' | 'PROPER' | 'PV' | 'QUARTILE' | 'QUOTIENT' | 'RADIANS' | 'RAND' | 'RANDBETWEEN' | 'RANK' | 'RATE' | 'RECEIVED' | 'REFTEXT' | 'REGISTER' | 'REGISTER.ID' | 'RELREF' | 'RENAME.COMMAND' | 'REPLACE' | 'REPLACEB' | 'REPT' | 'REQUEST' | 'RESET.TOOLBAR' | 'RESTART' | 'RESULT' | 'RESUME' | 'RETURN' | 'RIGHT' | 'RIGHTB' | 'ROMAN' | 'ROUND' | 'ROUNDBAHTDOWN' | 'ROUNDBAHTUP' | 'ROUNDDOWN' | 'ROUNDUP' | 'ROW' | 'ROWS' | 'RSQ' | 'RTD' | 'SAVE.DIALOG' | 'SAVE.TOOLBAR' | 'SCENARIO.GET' | 'SEARCH' | 'SEARCHB' | 'SECOND' | 'SELECTION' | 'SERIES' | 'SERIESSUM' | 'SET.NAME' | 'SET.VALUE' | 'SHOW.BAR' | 'SIGN' | 'SIN' | 'SINH' | 'SKEW' | 'SLN' | 'SLOPE' | 'SMALL' | 'SPELLING.CHECK' | 'SPREADBASE.DATA.FIELD' | 'SQRT' | 'SQRTPI' | 'STANDARDIZE' | 'STDEV' | 'STDEVA' | 'STDEVP' | 'STDEVPA' | 'STEP' | 'STEYX' | 'SUBSTITUTE' | 'SUBTOTAL' | 'SUM' | 'SUMIF' | 'SUMIFS' | 'SUMPRODUCT' | 'SUMSQ' | 'SUMX2MY2' | 'SUMX2PY2' | 'SUMXMY2' | 'SWITCH' | 'SYD' | 'T' | 'TAN' | 'TANH' | 'TBILLEQ' | 'TBILLPRICE' | 'TBILLYIELD' | 'TDIST' | 'TERMINATE' | 'TEXT' | 'TEXT.BOX' | 'TEXTJOIN' | 'TEXTREF' | 'THAIDAYOFWEEK' | 'THAIDIGIT' | 'THAIMONTHOFYEAR' | 'THAINUMSOUND' | 'THAINUMSTRING' | 'THAISTRINGLENGTH' | 'THAIYEAR' | 'TIME' | 'TIMEVALUE' | 'TINV' | 'TODAY' | 'TRANSPOSE' | 'TREND' | 'TRIM' | 'TRIMMEAN' | 'TRUE' | 'TRUNC' | 'TTEST' | 'TYPE' | 'UNREGISTER' | 'UPPER' | 'USDOLLAR' | 'VALUE' | 'VAR' | 'VARA' | 'VARP' | 'VARPA' | 'VDB' | 'VIEW.GET' | 'VLOOKUP' | 'VOLATILE' | 'WEEKDAY' | 'WEEKNUM' | 'WEIBULL' | 'WHILE' | 'WINDOW.TITLE' | 'WINDOWS' | 'WORKDAY' | 'XIRR' | 'XNPV' | 'YEAR' | 'YEARFRAC' | 'YIELD' | 'YIELDDISC' | 'YIELDMAT' | 'ZTEST') '(' WHITESPACES
        ;

FUTURE_FUNCTION_LIST
        : ('_xlfn.' ( 'ACOT' | 'ACOTH' | 'AGGREGATE' | 'ARABIC' | 'BASE' | 'BETA.DIST' | 'BETA.INV' | 'BINOM.DIST' | 'BINOM.DIST.RANGE' | 'BINOM.INV' | 'BITAND' | 'BITLSHIFT' | 'BITOR' | 'BITRSHIFT' | 'BITXOR' | 'BYCOL' | 'BYROW' | 'CEILING.MATH' | 'CEILING.PRECISE' | 'CHISQ.DIST' | 'CHISQ.DIST.RT' | 'CHISQ.INV' | 'CHISQ.INV.RT' | 'CHISQ.TEST' | 'COMBINA' | 'CONFIDENCE.NORM' | 'CONFIDENCE.T' | 'COT' | 'COTH' | 'COVARIANCE.P' | 'COVARIANCE.S' | 'CSC' | 'CSCH' | 'DAYS' | 'DECIMAL' | 'ERF.PRECISE' | 'ERFC.PRECISE' | 'EXPON.DIST' | 'F.DIST' | 'F.DIST.RT' | 'F.INV' | 'F.INV.RT' | 'F.TEST' | 'FIELDVALUE' | 'FILTERXML' | 'FLOOR.MATH' | 'FLOOR.PRECISE' | 'FORMULATEXT' | 'GAMMA' | 'GAMMA.DIST' | 'GAMMA.INV' | 'GAMMALN.PRECISE' | 'GAUSS' | 'HYPGEOM.DIST' | 'IFNA' | 'IMCOSH' | 'IMCOT' | 'IMCSC' | 'IMCSCH' | 'IMSEC' | 'IMSECH' | 'IMSINH' | 'IMTAN' | 'ISFORMULA' | 'ISOMITTED' | 'ISOWEEKNUM' | 'LAMBDA' | 'LET' | 'LOGNORM.DIST' | 'LOGNORM.INV' | 'MAKEARRAY' | 'MAP' | 'MODE.MULT' | 'MODE.SNGL' | 'MUNIT' | 'NEGBINOM.DIST' | 'NORM.DIST' | 'NORM.INV' | 'NORM.S.DIST' | 'NORM.S.INV' | 'NUMBERVALUE' | 'PDURATION' | 'PERCENTILE.EXC' | 'PERCENTILE.INC' | 'PERCENTRANK.EXC' | 'PERCENTRANK.INC' | 'PERMUTATIONA' | 'PHI' | 'POISSON.DIST' | 'QUARTILE.EXC' | 'QUARTILE.INC' | 'QUERYSTRING' | 'RANDARRAY' | 'RANK.AVG' | 'RANK.EQ' | 'REDUCE' | 'RRI' | 'SCAN' | 'SEC' | 'SECH' | 'SEQUENCE' | 'SHEET' | 'SHEETS' | 'SKEW.P' | 'SORTBY' | 'STDEV.P' | 'STDEV.S' | 'T.DIST' | 'T.DIST.2T' | 'T.DIST.RT' | 'T.INV' | 'T.INV.2T' | 'T.TEST' | 'UNICHAR' | 'UNICODE' | 'UNIQUE' | 'VAR.P' | 'VAR.S' | 'WEBSERVICE' | 'WEIBULL.DIST' | 'XLOOKUP' | 'XOR' | 'Z.TEST' | 'ECMA.CEILING' | 'ISO.CEILING' | 'NETWORKDAYS.INTL' | 'WORKDAY.INTL' | 'FORECAST.ETS' | 'FORECAST.ETS.CONFINT' | 'FORECAST.ETS.SEASONALITY' | 'FORECAST.LINEAR' | 'FORECAST.ETS.STAT' )) '(' WHITESPACES
        ;
*/

/*
 * Rather than split A1_REFERENCE, detect cell function through an extra token.
 * Since ANTLR takes the longest match, it will find this token, if it is a cell
 * function.
 */
CELL_FUNCTION_LIST
        : A1_CELL '(' WHITESPACES
        ;

/*
 * This is basically inlined `name_reference` along with open brace as a single
 * token. That removes a lot of ambiguity from atom rules to make grammar
 * LL(1).
 * ANTLR takes the longest match, so if name_reference is used as a UDF name,
 * this token will be matched. Must be last, so other functions have
 * precedence.
 */
USER_DEFINED_FUNCTION_NAME
        : NAME '(' WHITESPACES
        ;

/* --------------------------------- Name ---------------------------------- */

/*
 * Define NAME last, after other tokens, so ANTLR recognizes them as different tokens.
 *
 * A name MUST NOT have any of the following forms:
 * > TRUE or FALSE
 * > cell-reference
 * > function-list
 * > command-list
 * > future-function-list
 * > R1C1-cell-reference
 */
NAME
        : NAME_START_CHARACTER NAME_CHARACTERS?
        ;

fragment NAME_START_CHARACTER
        : UNDERSCORE
        | BACKSLASH
        | LETTER
        | NAME_BASE_CHARACTER
        ;

fragment UNDERSCORE
        : '_'
        ;

fragment BACKSLASH
        : '\\'
        ;

fragment LETTER
        : [a-zA-Z]
        ;

/*
 * Name should also be any code point that is a character, but I have trouble understanding the meaning.
 * Unicode 5.1 has valid codepoints for 0..10FFFF, but that includes surrogates, control, private ect...
 *
 * Original:
 * any code points which are characters as defined by the Unicode character properties, [UNICODE5.1] chapter 4
 * MUST NOT be 0x0-0x7F
 */
fragment NAME_BASE_CHARACTER
        : '\u0080' .. '\u{10FFFF}'
        ;

fragment NAME_CHARACTERS
        : NAME_CHARACTER+
        ;

fragment NAME_CHARACTER
        : NAME_START_CHARACTER
        | DECIMAL_DIGIT
        | FULL_STOP
        | QUESTIONMARK
        ;

fragment QUESTIONMARK
        : '?'
        ;

/* ---------------------------- Table reference ---------------------------- */

INTRA_TABLE_REFERENCE
        : SPACED_LBRACKET INNER_REFERENCE SPACED_RBRACKET
        | KEYWORD
        | '[' SIMPLE_COLUMN_NAME? ']'
        ;

fragment INNER_REFERENCE
        : KEYWORD_LIST
        | (KEYWORD_LIST SPACED_COMMA)? COLUMN_RANGE
        ;

fragment KEYWORD
        : '[#All]'
        | '[#Data]'
        | '[#Headers]'
        | '[#Totals]'
        | '[#This Row]'
        ;

fragment KEYWORD_LIST
        : KEYWORD
        | '[#Headers]' SPACED_COMMA '[#Data]'
        | '[#Data]' SPACED_COMMA '[#Totals]'
        ;

fragment COLUMN_RANGE
        : COLUMN (':' COLUMN)?
        ;

fragment COLUMN
        : SIMPLE_COLUMN_NAME
        | OPEN_SQUARE SPACE* SIMPLE_COLUMN_NAME SPACE* CLOSE_SQUARE
        ;

fragment SIMPLE_COLUMN_NAME
        : (ANY_NOSPACE_COLUMN_CHARACTER ANY_COLUMN_CHARACTER*)? ANY_NOSPACE_COLUMN_CHARACTER
        ;

fragment ESCAPE_COLUMN_CHARACTER
        : TICK
        | OPEN_SQUARE
        | CLOSE_SQUARE
        | '#'
        ;

fragment TICK
        : '\u0027'
        ;

fragment UNESCAPED_COLUMN_CHARACTER // MUST NOT match escape-column-character or space;
        : '\u0009'
        | '\u000A'
        | '\u000D'
        | '\u0021' .. '\u0022' // 0020 SPACE 0023 #
        | '\u0024' .. '\u0026' // 0027 '
        | '\u0028' .. '\u005A' // 005B [
        | '\u005C'             // 005D ]
        | '\u005E' .. '\uD7FF'
        | '\uE000' .. '\uFFFD'
        | '\u{10000}' .. '\u{10FFFF}' // ; as defined by the production Char in the [W3C-XML] section 2.2
        ;

fragment ANY_COLUMN_CHARACTER
        : ANY_NOSPACE_COLUMN_CHARACTER
        | SPACE
        ;

fragment ANY_NOSPACE_COLUMN_CHARACTER
        : UNESCAPED_COLUMN_CHARACTER
        | TICK ESCAPE_COLUMN_CHARACTER
        ;

fragment SPACED_COMMA
        : WHITESPACES COMMA WHITESPACES
        ;

fragment SPACED_LBRACKET
        : OPEN_SQUARE WHITESPACES
        ;

fragment SPACED_RBRACKET
        : WHITESPACES CLOSE_SQUARE
        ;

/* ---------------------------- Case Insensitive Fragments ---------------------------- */

fragment A : [Aa];
fragment B : [Bb];
fragment C : [Cc];
fragment D : [Dd];
fragment E : [Ee];
fragment F : [Ff];
fragment G : [Gg];
fragment H : [Hh];
fragment I : [Ii];
fragment J : [Jj];
fragment K : [Kk];
fragment L : [Ll];
fragment M : [Mm];
fragment N : [Nn];
fragment O : [Oo];
fragment P : [Pp];
fragment Q : [Qq];
fragment R : [Rr];
fragment S : [Ss];
fragment T : [Tt];
fragment U : [Uu];
fragment V : [Vv];
fragment W : [Ww];
fragment X : [Xx];
fragment Y : [Yy];
fragment Z : [Zz];