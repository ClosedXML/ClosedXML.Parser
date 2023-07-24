namespace ClosedXML.Parser.Ast;

public record StructureReferenceNode(string? Table, StructuredReferenceArea Area, string? FirstColumn, string? LastColumn) : AstNode;