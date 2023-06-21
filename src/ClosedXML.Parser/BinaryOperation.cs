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
    Plus,
    
    /// <summary><c>-</c></summary>
    Minus,

    /// <summary><c>*</c></summary>
    Mult,

    /// <summary><c>/</c></summary>
    Div,

    /// <summary><c>^</c></summary>
    Pow,

    /// <summary><c>,</c></summary>
    Union,
    
    /// <summary><c>SPACE</c></summary>
    Intersection,

    /// <summary><c>:</c></summary>
    Range,
}