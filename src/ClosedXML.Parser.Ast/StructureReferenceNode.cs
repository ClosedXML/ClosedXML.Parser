using System.Text;

namespace ClosedXML.Parser;

public record StructureReferenceNode(
    string? Table, 
    StructuredReferenceArea Area,
    string? FirstColumn,
    string? LastColumn) : AstNode
{
    public override string GetDisplayString()
    {
        var sb = new StringBuilder();
        if (Table is not null)
            sb.Append(Table);

        sb.Append('[');

        var list = new List<string>(3);
        if (Area != StructuredReferenceArea.None)
            list.Add(Area.GetDisplayString());

        if (FirstColumn is not null)
            list.Add($"[{FirstColumn}]");

        if (LastColumn is not null)
            list.Add($"[{LastColumn}]");

        sb.Append(string.Join(",", list));
        sb.Append(']');
        return sb.ToString();
    }
};