namespace ClosedXML.Parser.Pratt;

/// <summary>
/// Various methods that are not present in .net standard 2.0.
/// </summary>
internal static class CompatUtils
{
    /// <summary>
    /// Replacement for <c>char.IsAsciiLetter</c> that isn't in the netstandard 2.0
    /// </summary>
    public static bool IsAsciiLetter(char c)
    {
        return c is >= 'A' and <= 'Z' ||
               c is >= 'a' and <= 'z';
    }

    /// <summary>
    /// Replacement for <c>char.IsAsciiDigit</c> that isn't in the netstandard 2.0
    /// </summary>
    public static bool IsAsciiDigit(char c)
    {
        return c is >= '0' and <= '9';
    }
}
