using System;
using Newtonsoft.Json;

namespace ClosedXML.Parser.Function;

internal class AstNodeConverter : JsonConverter<AstNode>
{
    private readonly ReferenceStyle _style;

    internal AstNodeConverter(ReferenceStyle style)
    {
        _style = style;
    }

    public override AstNode ReadJson(JsonReader reader, Type objectType, AstNode existingValue, bool hasExistingValue, JsonSerializer serializer)
    {
        throw new NotSupportedException("Deserialization of AST is not supported.");
    }

    public override void WriteJson(JsonWriter writer, AstNode value, JsonSerializer serializer)
    {
        WriteNode(writer, value);
    }

    private void WriteNode(JsonWriter writer, AstNode value)
    {
        writer.WriteStartObject();
        writer.WritePropertyName("type");
        writer.WriteValue(value.GetTypeString());

        writer.WritePropertyName("content");
        writer.WriteValue(value.GetDisplayString(_style));

        writer.WritePropertyName("children");
        writer.WriteStartArray();
        foreach (var child in value.Children)
            WriteNode(writer, child);

        writer.WriteEndArray();
        writer.WriteEndObject();
    }
}