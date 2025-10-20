namespace ClosedXML.Parser.Pratt;

/// <summary>
/// An info about node used during parsing.
/// </summary>
/// <typeparam name="T">The <c>TNode</c> type of a node from <see cref="IAstFactory{TScalarValue,TNode,TContext}"/>.</typeparam>
internal readonly struct Node<T>
{
    public Node(T value, int start, int end)
        : this(value, new SymbolRange(start, end))
    {
    }

    public Node(T value, SymbolRange range)
    {
        Value = value;
        Range = range;
    }

    /// <summary>
    /// Parsed value of a node, created by the <see cref="IAstFactory{TScalarValue,TNode,TContext}"/>.
    /// </summary>
    public T Value { get; }

    /// <summary>
    /// A range that was used to created the node.
    /// </summary>
    public SymbolRange Range { get; }

    public static implicit operator T(Node<T> node)
    {
        return node.Value;
    }

    internal Node<T> ExtendLeft(Token token)
    {
        return new Node<T>(Value, token.Range.ExtendRight(Range));
    }

    internal Node<T> ExtendRight(Token token)
    {
        return new Node<T>(Value, Range.ExtendRight(token.Range));
    }
}
