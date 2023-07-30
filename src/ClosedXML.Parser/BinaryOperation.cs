namespace ClosedXML.Parser;

public enum BinaryOperation
{
    /// <summary><c>&amp;</c></summary>
    Concat,

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

    /// <summary><c>,</c></summary>
    Union,

    /// <summary><c>SPACE</c></summary>
    Intersection,

    /// <summary><c>:</c></summary>
    Range,
}