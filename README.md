# ClosedParser

ClosedParser is a project to parse OOXML grammar to create an abstract syntax tree that can be later evaluated.

Official source for the grammar is [MS-XML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e), chapter 2.2.2 Formulas. The provided grammar is not usable for parser generators, it's full of ambiguities and the rules don't take into account operator precedence.

# Goals

* __Performance__ - ClosedXML needs to parse formula really fast. Limit allocation and so on.
* __Evaluation oriented__ - Parser should concentrates on creation of abstract syntax trees, not concrete syntax tree. Goal is evaluation of formulas, not transformation.
* __Multi-use__ - Formulas are mostly used in cells, but there are other places with different grammar rules (e.g. sparklines, data validation)
* __Multi notation (A1 or R1C1)__ - Parser should be able to parse both A1 and R1C1 formulas. I.e. `SUM(R5)` can mean return sum of cell `R5` in _A1_ notation, but return sum of all cells on row 5 in _R1C1_ notation.

Project uses ANTLR4 grammar file as the source of truth and a lexer. There is also ANTLR parser is not used, but is used as a basis of recursive descent parser (ANTLR takes up 8 seconds vs RDS 700ms for parsing of enron dataset).

ANTLR4 one of few maintained parser generators with C# target.

The project has a low priority, XLParser mostly works, but in the long term, replacement is likely.

# Why not use XLParser

ClosedXML is currently using [XLParser](https://github.com/spreadsheetlab/XLParser) and transforming the concrete syntax tree to abstract syntax tree.

* Speed:
  * Grammar extensively uses regexps extensively. Regexs are slow, especially for NET4x target, allocates extra memory. XLParser takes up _47_ seconds for Enron dataset on .NET Framework. .NET teams had made massive improvements on regexs, so it takes only _16_ seconds on NET7.
  * IronParser needs to determine all possible tokens after every token, that is problematic, even with the help of `prefix` hints.
* AST: XLParser creates concentrates on creation of concrete syntax tree, but for ClosedXML, we need abstract syntax tree for evaluation. IronParser is not very friendly in that regard
* ~~XLParser uses `IronParser`, an unmaintained project~~ (IronParser recently released version 1.2).
* Doesn't have support for lambdas and R1C1 style.

ANTLR lexer takes up about 3.2 seconds for Enron dataset. With ANTLR parsing, it takes up 11 seconds. I want that 7+ seconds in performance and no allocation, so RDS that takes up 700 ms.

## Current state

ENRON dataset parsing RDS + ANTLR lexer

* Total: *946320*
* Elapsed: *2964 ms*
* Per formula: *3.132 μs*

3μs per formula should be something like 9000 instructions (under unrealistic assumption 1 instruction per 1 Hz), so not much space to improve.

Bad ones fail on external workbook name reference that uses apostrophe (e.g `[1]!'Some name in external wb'`), otherwise they are parsed.

## Debugging

Use [vscode-antlr4](https://github.com/mike-lischke/vscode-antlr4/blob/master/doc/grammar-debugging.md) plugin for debugging the grammar.

## Resource

* [MS-XML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e)
* [Simplified XLParser grammar](https://github.com/spreadsheetlab/XLParser/blob/master/doc/ebnf.pdf) and [tokens](https://github.com/spreadsheetlab/XLParser/blob/master/doc/tokens.pdf).
* [Getting Started With ANTLR in C#](https://tomassetti.me/getting-started-with-antlr-in-csharp/)
