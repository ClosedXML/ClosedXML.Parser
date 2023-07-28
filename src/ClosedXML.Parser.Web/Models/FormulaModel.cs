namespace ClosedXML.Parser.Web.Models;

public class FormulaModel
{
    /// <summary>
    /// Style of cell references in formula. Either <c>A1</c> or <c>R1C1</c>.
    /// </summary>
    public ReferenceStyle Style { get; set; }

    /// <summary>
    /// Text of parsed formula.
    /// </summary>
    public string Formula { get; set; }

    /// <summary>
    /// The root node of AST, if formula is parseable. Null, otherwise.
    /// </summary>
    public AstNode? Ast { get; set; }

    /// <summary>
    /// If formula is unparsable, the error of why it is unparsable.
    /// </summary>
    public string? Error { get; set; }
}