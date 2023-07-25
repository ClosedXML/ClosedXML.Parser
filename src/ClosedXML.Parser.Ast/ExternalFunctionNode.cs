namespace ClosedXML.Parser;

public record ExternalFunctionNode(int WorkbookIndex, string? Sheet, string Name) : AstNode;