using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Parser.Rolex;

/// <summary>
/// A class required by a Rolex tool. Never used.
/// </summary>
internal class TableTokenizer
{
    [SuppressMessage("ReSharper", "UnusedParameter.Local", Justification = "Parameters must be present to satisfy machine generated contract.")]
    protected TableTokenizer(DfaEntry[] dfaTable, int[][] blockEnds, int[] nodeFlags, IEnumerable<char> input)
    {
        throw new NotSupportedException("This class should never be instantiated.");
    }
}