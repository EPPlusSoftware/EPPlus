using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal class ChainingContextualHandler : IGsubLookupHandler
    {
        public ushort LookupType => 6;

        public void Discover(FontSubsettingContext context, LookupTable lookup, GsubSubsetProcessor processor)
        {
            foreach (var subTable in lookup.SubTables.OfType<ChainingContextualSubstFormat3>())
            {
                // 1. Check if we have a match on the first character (e.g., 'f')
                // Using the name InputCoverage here as it's most common in format 3
                if (subTable.InputCoverages != null && subTable.InputCoverages.Count > 0)
                {
                    var initialCoverage = subTable.InputCoverages[0];

                    if (AnyGlyphInSubset(initialCoverage, context.IncludedGlyphs))
                    {
                        // 2. Loop through records to find linked lookups
                        foreach (var record in subTable.SubstLookupRecords)
                        {
                            int lookupIndex = record.LookupListIndex;

                            // 3. Get the actual LookupTable instance from the original font
                            // We go through context.OriginalFont to find the correct table in the list
                            if (lookupIndex >= 0 && lookupIndex < context.OriginalFont.GsubTable.LookupList.Lookups.Count)
                            {
                                var targetLookup = context.OriginalFont.GsubTable.LookupList.Lookups[lookupIndex];

                                // 4. Now the argument matches! (context, LookupTable)
                                processor.DiscoverLookup(context, targetLookup);
                            }
                        }
                    }
                }
            }
        }

        private bool AnyGlyphInSubset(CoverageTable coverage, HashSet<ushort> includedGlyphs)
        {
            if (coverage == null) return false;
            // Get all GIDs that this coverage table includes
            var coveredGids = coverage.GetCoveredGlyphs();
            // Check if any of these GIDs exist in our current subset
            return coveredGids.Any(gid => includedGlyphs.Contains(gid));
        }
        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newLookup = new LookupTable
            {
                LookupType = 6,
                LookupFlag = oldLookup.LookupFlag,
                SubTables = new List<FontTableElement>()
            };

            foreach (var subtable in oldLookup.SubTables.OfType<ChainingContextualSubstFormat3>())
            {
                var rewritten = subtable.Rewrite(context);

                if (rewritten != null)
                {
                    newLookup.SubTables.Add(rewritten);
                }
            }

            if (newLookup.SubTables.Count > 0)
            {
                return newLookup;
            }

            return null;
        }
    }
}