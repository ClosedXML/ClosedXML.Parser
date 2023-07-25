namespace ClosedXML.Parser;

public record StructureReferenceNode(string? Table, StructuredReferenceArea Area, string? FirstColumn, string? LastColumn) : AstNode;