using System.Text;

namespace ClosedXML.Parser;

public record ArrayNode(int Rows, int Columns, IReadOnlyList<ScalarValue> Elements) : AstNode
{
    public override string GetDisplayString()
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
        unchecked
        {
            int hashCode = base.GetHashCode();
            hashCode = (hashCode * 397) ^ Rows;
            hashCode = (hashCode * 397) ^ Columns;
            return hashCode;
        }
    }
}