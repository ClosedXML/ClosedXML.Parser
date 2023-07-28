using System.Text;

namespace ClosedXML.Parser;

public record ArrayNode(int Rows, int Columns, IReadOnlyList<ScalarValue> Elements) : AstNode
{
    public override string GetDisplayString(ReferenceStyle style)
    {
        var sb = new StringBuilder();

        var idx = 0;
        sb.Append('{');
        sb.Append(Elements[idx++].GetDisplayString());
        for (var col = 1; col < Columns; ++col)
            sb.Append(',').Append(Elements[idx++].GetDisplayString());

        for (var row = 1; row < Rows; ++row)
        {
            sb.Append(';');
            sb.Append(Elements[idx++].GetDisplayString());
            for (var col = 1; col < Columns; ++col)
                sb.Append(',').Append(Elements[idx++].GetDisplayString());
        }
        sb.Append('}');
        
        return sb.ToString();
    }

    public virtual bool Equals(ArrayNode? other)
    {
        if (ReferenceEquals(null, other))
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return base.Equals(other) &&
               Rows == other.Rows &&
               Columns == other.Columns &&
               Elements.SequenceEqual(other.Elements);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Rows, Columns);
    }
}