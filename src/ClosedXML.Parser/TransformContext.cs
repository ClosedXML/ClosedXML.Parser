namespace ClosedXML.Parser;

internal sealed class TransformContext
{
    public TransformContext(string formula, int row, int col)
    {
        Formula = formula;
        Row = row;
        Col = col;
    }

    public string Formula { get; }
    
    public int Row { get; } 
    
    public int Col { get; }
}