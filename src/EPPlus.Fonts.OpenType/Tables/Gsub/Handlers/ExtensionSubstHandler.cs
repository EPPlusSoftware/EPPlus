using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub.Handlers;
using System.Collections.Generic;

/// <summary>
/// Handler for GSUB Lookup Type 7: Extension Substitution.
/// <para>
/// The loader keeps each wrapped subtable in an <see cref="ExtensionSubstSubTable"/> together
/// with its inner lookup type. This handler unwraps them into a lookup of the inner type and
/// delegates to the handler for that type.
/// </para>
/// <para>
/// The rewritten lookup is returned with the inner lookup type, not re-wrapped as type 7.
/// The subset is written from scratch and is small, so there is no 16 bit offset overflow
/// to avoid. This matches how <c>ExtensionPosHandler</c> handles GPOS type 9.
/// </para>
/// </summary>
internal class ExtensionSubstHandler : IGsubLookupHandler
{
    public ushort LookupType => 7;

    public void Discover(FontSubsettingContext context, LookupTable lookup, GsubSubsetProcessor processor)
    {
        LookupTable unwrapped = Unwrap(lookup);
        if (unwrapped != null)
        {
            processor.DiscoverLookup(context, unwrapped);
        }
    }

    public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
    {
        LookupTable unwrapped = Unwrap(oldLookup);
        if (unwrapped == null) return null;

        IGsubLookupHandler handler = context.GsubProcessor.GetHandler(unwrapped.LookupType);
        if (handler == null) return null;

        // The inner handlers set LookupType to their own type, so the result already
        // carries the inner type.
        LookupTable rewritten = handler.Rewrite(context, unwrapped);
        if (rewritten == null || rewritten.SubTables.Count == 0) return null;

        return rewritten;
    }

    /// <summary>
    /// Builds a lookup of the inner lookup type from the wrapped subtables.
    /// Returns null if no wrapped subtables were loaded.
    /// </summary>
    private static LookupTable Unwrap(LookupTable lookup)
    {
        if (lookup.SubTables == null || lookup.SubTables.Count == 0) return null;

        ushort innerLookupType = 0;
        List<FontTableElement> innerSubTables = new List<FontTableElement>();

        foreach (FontTableElement subTable in lookup.SubTables)
        {
            ExtensionSubstSubTable extension = subTable as ExtensionSubstSubTable;
            if (extension == null || extension.ExtendedSubTable == null) continue;

            if (innerLookupType == 0)
            {
                innerLookupType = extension.ExtensionLookupType;
            }
            else if (extension.ExtensionLookupType != innerLookupType)
            {
                // All subtables of a lookup must share the same type per the OpenType
                // specification. Ignore mismatching ones rather than emitting a lookup
                // whose type does not describe its subtables.
                continue;
            }

            innerSubTables.Add(extension.ExtendedSubTable);
        }

        // Type 7 can never be the inner type; guard against recursion on malformed fonts.
        if (innerLookupType == 0 || innerLookupType == 7) return null;

        return new LookupTable
        {
            LookupType = innerLookupType,
            LookupFlag = lookup.LookupFlag,
            MarkFilteringSet = lookup.MarkFilteringSet,
            SubTables = innerSubTables
        };
    }
}