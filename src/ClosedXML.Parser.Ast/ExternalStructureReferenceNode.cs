namespace ClosedXML.Parser;

public record ExternalStructureReferenceNode(int WorkbookIndex, string Table, StructuredReferenceArea Area, string? FirstColumn, string? LastColumn) : AstNode;