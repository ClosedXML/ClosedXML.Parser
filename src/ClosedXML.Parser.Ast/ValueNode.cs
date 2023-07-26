namespace ClosedXML.Parser;

public record ValueNode(string Type, object Value) : AstNode
{
    public ValueNode(double value) : this("Number", value) { }
    public ValueNode(bool value) : this("Logical", value) { }

    public override string GetTypeString() => Type;

    public override string GetDisplayString()
    {
        return Value?.ToString() ?? "BLANK";
    }
};