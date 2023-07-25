namespace ClosedXML.Parser;

public readonly record struct ScalarValue(string Type, object Value)
{
    public ScalarValue(double value) : this("Number", value) { }

    public ScalarValue(bool value) : this("Logical", value) { }

    public ScalarValue(string value) : this("Text", value) { }
};