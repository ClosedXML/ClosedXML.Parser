using System.Text.Json;
using System.Text.Json.Serialization;

namespace ClosedXML.Parser.Web;

internal class AstNodeConverter : JsonConverter<AstNode>
{
    private readonly ReferenceStyle _style;

    internal AstNodeConverter(ReferenceStyle style)
    {
        _style = style;
    }

    public override AstNode? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        throw new NotSupportedException("Deserialization of AST is not supported.");
    }

    public override void Write(Utf8JsonWriter writer, AstNode value, JsonSerializerOptions options)
    {
        WriteNode(writer, value, options);
    }

    private void WriteNode(Utf8JsonWriter writer, AstNode value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        WritePropertyName("type");
        writer.WriteStringValue(value.GetTypeString());

        WritePropertyName("content");
        writer.WriteStringValue(value.GetDisplayString(_style));

        WritePropertyName("children");
        writer.WriteStartArray();
        foreach (var child in value.Children)
            WriteNode(writer, child, options);

        writer.WriteEndArray();
        writer.WriteEndObject();

        void WritePropertyName(string propertyName)
        {
            var normalizedName = options.PropertyNamingPolicy?.ConvertName(propertyName) ?? propertyName;
            writer.WritePropertyName(normalizedName);
        }
    }
}