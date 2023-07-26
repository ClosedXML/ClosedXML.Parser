using System.Text.Json;
using System.Text.Json.Serialization;

namespace ClosedXML.Parser.Web;

public class AstNodeConverter : JsonConverter<AstNode>
{
    public override AstNode? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        throw new NotSupportedException("Deserialization of AST is not supported.");
    }

    public override void Write(Utf8JsonWriter writer, AstNode value, JsonSerializerOptions options)
    {
        WriteNode(writer, value, options);
    }

    private static void WriteNode(Utf8JsonWriter writer, AstNode value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        WritePropertyName("type");
        writer.WriteStringValue(value.GetTypeString());

        WritePropertyName("content");
        writer.WriteStringValue(value.GetDisplayString());

        WritePropertyName("children");
        writer.WriteStartArray();
        foreach (var child in value.Children)
        {
            WriteNode(writer, child, options);
        }
        writer.WriteEndArray();
        writer.WriteEndObject();

        void WritePropertyName(string propertyName)
        {
            var normalizedName = options.PropertyNamingPolicy?.ConvertName(propertyName) ?? propertyName;
            writer.WritePropertyName(normalizedName);
        }
    }
}