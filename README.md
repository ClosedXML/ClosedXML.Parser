# ClosedParser

ClosedParser is a project to parse OOXML grammar.

Official source for the grammar is [MS-XML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e), chapter 2.2.2 Formulas. The provided grammar is not usable for parser generators, it's full of ambiguities and the rules don't take into account operator precedence.

# Goals

Fast parser to create AST for cell formulas and other places that need to parse formula (e.g. sparklines, data validation) for ClosedXML. There shoudl be a ANTLR4 grammar file and possibly in the future hand-made recrusive descent parser that doesn't construct concrete syntax tree at all, but directly produces AST.

ANTLR4 one of few maintained parser generators with C# target.

The project has a low priority, XLParser mostly works, but in the long term, replacement is likely.

# Why not use XLParser

ClosedXML is currently using [XLParser](https://github.com/spreadsheetlab/XLParser) and transforming the concrete syntax tree to abstract syntax tree.

* Speed:
  * Grammar extensively uses regexps extensively. Regexs are slow, especially for NET4x target, allocates extra memory.
  * IronParser needs to determine all possible tokens after every token, that is problematic, even with the help of `prefix` hints.
* AST: XLParser creates concentrates on creation of concrete syntax tree, but for ClosedXML, we need abstract syntax tree for evaluation. IronParser is not very friendly in that regard
* ~~XLParser uses `IronParser`, an unmaintained project~~ (IronParser recently released version 1.2).
* Doesn't have support for lambdas and R1C1 style.

ANTLR lexer takes up about 3.2 seconds for Enron dataset. With parsing, it takes up 11 seconds. I want that 7+ seconds in performance and no allocation.

## Debugging

Use [vscode-antlr4](https://github.com/mike-lischke/vscode-antlr4/blob/master/doc/grammar-debugging.md) plugin for debugging the grammar.

Current state - ENRON dataset parsing
> Total: 946320
> Good Count: 945667
> Bad Count: 653
> Elapsed: 9657 ms

Bad ones fail on external workbook name reference that uses apostrophe (e.g `[1]!'Some name in external wb'`), otherwise they are parsed.

* Solve cell formulas
* Convert to LL(1) parser
* Prerequisite for recursive descent parser
  * Remove left recusion
  * Remove left factoring
* Make recursive descent parser

## Resource

* [MS-XML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e)
* [Simplified XLParser grammar](https://github.com/spreadsheetlab/XLParser/blob/master/doc/ebnf.pdf) and [tokens](https://github.com/spreadsheetlab/XLParser/blob/master/doc/tokens.pdf).
* [Getting Started With ANTLR in C#](https://tomassetti.me/getting-started-with-antlr-in-csharp/)
