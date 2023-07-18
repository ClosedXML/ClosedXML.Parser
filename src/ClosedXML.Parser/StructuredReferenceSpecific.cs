using System;

namespace ClosedXML.Parser;

[Flags]
public enum StructuredReferenceSpecific
{
    None = 0,

    /// <summary>
    /// <c>[#Data]</c> - only data cells of a table, without headers or totals.
    /// </summary>
    Data = 1 << 1,

    /// <summary>
    /// <c>[#Headers]</c> - only header rows of a table, if it exists. If there isn't header row, <c>#REF!</c>.
    /// </summary>
    Headers = 1 << 2,

    /// <summary>
    /// <c>[#Totals]</c> - only totals rows of a table, if it exists. If there isn't totals row, <c>#REF!</c>.
    /// </summary>
    Totals = 1 << 3,

    /// <summary>
    /// <c>[#All]</c> - all cells of a table. 
    /// </summary>
    All = Data | Headers | Totals,

    /// <summary>
    /// <c>[#This Row]</c> - only the same data row as the referencing cell. <c>#VALUE!</c> if not on a data row
    /// (e.g. headers or totals) or out of a table.
    /// </summary>
    ThisRow = 1 << 4
}