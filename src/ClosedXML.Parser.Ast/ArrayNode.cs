namespace ClosedXML.Parser.Ast;

public record ArrayNode(int Rows, int Columns, IReadOnlyList<ScalarValue> Elements) : AstNode
{
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