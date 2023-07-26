namespace ClosedXML.Parser;

public abstract record AstNode
{
    public AstNode[] Children { get; init; } = Array.Empty<AstNode>();

    public abstract string GetDisplayString();

    public virtual bool Equals(AstNode? other) => other is not null && Children.SequenceEqual(other.Children);

    public override int GetHashCode() => Children.Sum(child => child.GetHashCode());
}