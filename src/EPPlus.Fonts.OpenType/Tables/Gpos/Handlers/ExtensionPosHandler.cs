/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS lookup handler interface
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using EPPlus.Fonts.OpenType.Tables.Gpos.Handlers;
using System.Collections.Generic;

internal class ExtensionPosHandler : IGposLookupHandler
{
    public ushort LookupType => 9;

    public void Discover(FontSubsettingContext context, LookupTable lookup, GposSubsetProcessor processor)
    {
        // Extension wraps another lookup - unwrap and delegate
        foreach (var subtable in lookup.SubTables)
        {
            if (subtable is MarkToBaseSubTableFormat1)
            {
                // It's a wrapped MarkToBase - delegate to MarkToBase handler
                var wrappedLookup = new LookupTable
                {
                    LookupType = 4,
                    SubTables = new List<FontTableElement> { subtable }
                };

                var handler = processor.GetHandler(4);
                if (handler != null)
                {
                    handler.Discover(context, wrappedLookup, processor);
                }
            }
            // Add cases for other wrapped types as needed
        }
    }

    public LookupTable Rewrite(FontSubsettingContext context, LookupTable lookup)
    {
        // Unwrap and delegate to the actual handler
        if (lookup.SubTables.Count == 0)
            return null;

        var firstSubtable = lookup.SubTables[0];

        if (firstSubtable is MarkToBaseSubTableFormat1)
        {
            // Create a temporary Type 4 lookup
            var wrappedLookup = new LookupTable
            {
                LookupType = 4,
                LookupFlag = lookup.LookupFlag,
                SubTables = lookup.SubTables
            };

            var handler = context.GposProcessor.GetHandler(4);
            if (handler != null)
            {
                var rewritten = handler.Rewrite(context, wrappedLookup);
                if (rewritten != null && rewritten.SubTables.Count > 0)
                {
                    // Wrap back in Extension
                    return new LookupTable
                    {
                        LookupType = 9,
                        LookupFlag = lookup.LookupFlag,
                        SubTables = rewritten.SubTables
                    };
                }
            }
        }

        return null;
    }
}