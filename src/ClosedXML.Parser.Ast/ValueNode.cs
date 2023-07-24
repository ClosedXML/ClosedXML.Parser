namespace ClosedXML.Parser.Ast;

public record ValueNode(string Type, object Value) : AstNode
{
    public ValueNode(double value) : this("Number", value) { }
    public ValueNode(bool value) : this("Logical", value) { }
};