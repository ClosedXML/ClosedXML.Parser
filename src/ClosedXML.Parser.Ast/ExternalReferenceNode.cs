namespace ClosedXML.Parser;

public record ExternalReferenceNode(int WorkbookIndex, CellArea Reference) : AstNode;