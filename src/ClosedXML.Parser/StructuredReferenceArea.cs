using System;

namespace ClosedXML.Parser;

/// <summary>
/// Structure reference is basically a set of cells in an area of an intersection between a range of columns
/// in a table and a vertical range. This enum represents possible values. Thanks to the pattern of a structure
/// reference token, the vertical range of a formula is always continuous (i.e. no <c>Headers</c> and <c>Totals</c>
/// together).
/// </summary>
/// <remarks>
/// The documentation calls it *Item specifier* and grammar *keywords*. Both rather unintuitive names, so *area*
/// is used instead.
/// </remarks>
[Flags]
public enum StructuredReferenceArea
{
    /// <summary>
    /// Nothing was specified in the structure reference. Should have same impact as <see cref="Data"/>.
    /// </summary>
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