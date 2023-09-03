namespace ClosedXML.Parser;

/// <summary>
/// Binary operations that can occur in the formula.
/// </summary>
public enum BinaryOperation
{
    #region Text operators
    
    /// <summary><c>&amp;</c></summary>
    Concat,

    #endregion

    #region Arithmetic operators

    /// <summary><c>+</c></summary>
    Addition,

    /// <summary><c>-</c></summary>
    Subtraction,

    /// <summary><c>*</c></summary>
    Multiplication,

    /// <summary><c>/</c></summary>
    Division,

    /// <summary><c>^</c></summary>
    Power,

    #endregion

    #region Comparison operators

    /// <summary><c>&gt;=</c></summary>
    GreaterOrEqualThan,

    /// <summary><c>&lt;=</c></summary>
    LessOrEqualThan,

    /// <summary><c>&lt;</c></summary>
    LessThan,

    /// <summary><c>&gt;</c></summary>
    GreaterThan,

    /// <summary><c>!=</c></summary>
    NotEqual,

    /// <summary><c>=</c></summary>
    Equal,

    #endregion

    #region Range operators

    /// <summary><c>,</c></summary>
    Union,

    /// <summary><c>SPACE</c></summary>
    Intersection,

    /// <summary><c>:</c></summary>
    Range,

    #endregion
}