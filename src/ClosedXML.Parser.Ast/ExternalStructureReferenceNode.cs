namespace ClosedXML.Parser.Ast;

public record ExternalStructureReferenceNode(int WorkbookIndex, string Table, StructuredReferenceArea Area, string? FirstColumn, string? LastColumn) : AstNode;