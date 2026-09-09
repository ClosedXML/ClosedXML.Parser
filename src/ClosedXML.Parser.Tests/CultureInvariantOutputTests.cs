using System.Globalization;

namespace ClosedXML.Parser.Tests;

/// <summary>
/// R1C1 is a machine format, so the writer must emit the ASCII hyphen-minus for negative
/// relative offsets. Cultures whose <c>NumberFormat.NegativeSign</c> is the Unicode MINUS
/// SIGN (U+2212) would otherwise produce text the R1C1 reader cannot parse back.
/// </summary>
public class CultureInvariantOutputTests
{
    private const string UnicodeMinus = "−";
    private const string AsciiHyphen = "-";

    /// <summary>
    /// Culture and the negative sign it is pinned to. The sign is set explicitly rather than
    /// taken from the platform, because the same culture reports a different one per
    /// globalization mode - ICU gives sv-SE the U+2212 that triggers the bug, while NLS
    /// (<c>DOTNET_SYSTEM_GLOBALIZATION_USENLS=1</c>) gives it an ASCII hyphen. Reading the
    /// sign from the host would let these tests pass without exercising the bug at all.
    /// </summary>
    public static TheoryData<string, string> Cultures => new()
    {
        { "en-US", AsciiHyphen },
        { "sv-SE", UnicodeMinus },
        { "fi-FI", UnicodeMinus },
        { "nb-NO", UnicodeMinus },
        { "cs-CZ", AsciiHyphen },
    };

    [Theory]
    [MemberData(nameof(Cultures))]
    public void ToR1C1_writes_negative_offsets_with_ascii_hyphen(string cultureName, string negativeSign)
    {
        using var _ = new CultureScope(cultureName, negativeSign);

        Assert.Equal("RC[-2]/RC[-1]", FormulaConverter.ToR1C1("A1/B1", 1, 3));
        Assert.Equal("R[-1]C", FormulaConverter.ToR1C1("A1", 2, 1));
    }

    [Theory]
    [MemberData(nameof(Cultures))]
    public void ToR1C1_output_round_trips_back_to_A1(string cultureName, string negativeSign)
    {
        using var _ = new CultureScope(cultureName, negativeSign);

        var r1c1 = FormulaConverter.ToR1C1("A1/B1", 1, 3);

        Assert.Equal("A1/B1", FormulaConverter.ToA1(r1c1, 1, 3));
    }

    [Theory]
    [MemberData(nameof(Cultures))]
    public void GetDisplayStringR1C1_uses_ascii_hyphen(string cultureName, string negativeSign)
    {
        using var _ = new CultureScope(cultureName, negativeSign);

        var rowCol = new RowCol(ReferenceAxisType.Relative, -3, ReferenceAxisType.Relative, -7, R1C1);

        Assert.Equal("R[-3]C[-7]", rowCol.GetDisplayStringR1C1());
    }

    /// <summary>
    /// Makes <see cref="CultureInfo.CurrentCulture"/> a copy of <paramref name="name"/> with a
    /// known negative sign, so the test outcome doesn't depend on the host's globalization data.
    /// </summary>
    private sealed class CultureScope : IDisposable
    {
        private readonly CultureInfo _original = CultureInfo.CurrentCulture;

        public CultureScope(string name, string negativeSign)
        {
            var culture = (CultureInfo)CultureInfo.GetCultureInfo(name).Clone();
            culture.NumberFormat.NegativeSign = negativeSign;

            // Guard against a future .NET making the clone's NumberFormat read-only, which would
            // silently drop the sign and leave the negative-sign cultures no longer covered.
            Assert.Equal(negativeSign, culture.NumberFormat.NegativeSign);

            CultureInfo.CurrentCulture = culture;
        }

        public void Dispose() => CultureInfo.CurrentCulture = _original;
    }
}
