namespace ClosedXML.Parser.Ast;

public record ExternalReferenceNode(int WorkbookIndex, CellArea Reference) : AstNode;