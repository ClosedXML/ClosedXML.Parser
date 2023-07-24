namespace ClosedXML.Parser.Ast;

public record ExternalFunctionNode(int WorkbookIndex, string? Sheet, string Name) : AstNode;