namespace ClosedXML.Parser.Pratt;

/// <summary>
/// Values of binding power for operators in an expression. Higher number = higher binding power.
/// Precedence of operators is specified by ISO-29500:18.17.2.2. Operators that have the same
/// precedence associate left-to-right.
/// </summary>
internal static class BindingPower
{
    internal const int Addition = 3;
    internal const int Subtraction = 3;
    internal const int Multiplication = 4;
    internal const int Division = 4;
    internal const int Exponentiation = 5;
}
