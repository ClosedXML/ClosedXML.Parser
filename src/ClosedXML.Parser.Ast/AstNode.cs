namespace ClosedXML.Parser;

public abstract record AstNode
{
    public AstNode[] Children { get; init; } = Array.Empty<AstNode>();

    /// <summary>
    /// Render node and its children in a reference style.
    /// </summary>
    public abstract string GetDisplayString(ReferenceStyle style);

    public virtual string GetTypeString() => GetType().Name[..^4]; // Strip Node suffix

    public virtual bool Equals(AstNode? other) => other is not null && Children.SequenceEqual(other.Children);

    public override int GetHashCode() => Children.Sum(child => child.GetHashCode());
}