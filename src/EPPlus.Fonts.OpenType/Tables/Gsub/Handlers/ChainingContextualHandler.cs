using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Coverage;
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
                // 1. Kontrollera om vi har en match på första tecknet (t.ex. 'f')
                // Jag använder namnet InputCoverage här då det är vanligast i format 3
                if (subTable.InputCoverages != null && subTable.InputCoverages.Count > 0)
                {
                    var initialCoverage = subTable.InputCoverages[0];

                    if (AnyGlyphInSubset(initialCoverage, context.IncludedGlyphs))
                    {
                        // 2. Loopa igenom records för att hitta länkade lookups
                        foreach (var record in subTable.SubstLookupRecords)
                        {
                            int lookupIndex = record.LookupListIndex;

                            // 3. Hämta den faktiska LookupTable-instansen från den ursprungliga fonten
                            // Vi går via context.OriginalFont för att hitta rätt tabell i listan
                            if (lookupIndex >= 0 && lookupIndex < context.OriginalFont.GsubTable.LookupList.Lookups.Count)
                            {
                                var targetLookup = context.OriginalFont.GsubTable.LookupList.Lookups[lookupIndex];

                                // 4. Nu matchar argumentet! (context, LookupTable)
                                processor.DiscoverLookup(context, targetLookup);

                                System.Diagnostics.Debug.WriteLine($"Chaining: Följer länk till Lookup {lookupIndex} (Type {targetLookup.LookupType})");
                            }
                        }
                    }
                }
            }
        }

        private bool AnyGlyphInSubset(CoverageTable coverage, HashSet<ushort> includedGlyphs)
        {
            if (coverage == null) return false;
            // Hämta alla GID som denna täckningstabell omfattar
            var coveredGids = coverage.GetCoveredGlyphs();
            // Kolla om något av dessa GID finns i vårt nuvarande subset
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
                // Här anropas subtabellens egen Rewrite som vi precis jobbade med
                var rewritten = subtable.Rewrite(context);

                if (rewritten != null)
                {
                    newLookup.SubTables.Add(rewritten);
                    // HÄR loggar vi!
                    System.Diagnostics.Debug.WriteLine($"Chaining Rewrite: Behåller en subtabell med {rewritten.InputCoverages.Count} input-positioner.");
                }
            }

            if (newLookup.SubTables.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"Chaining Rewrite: Lookup färdig. Totalt {newLookup.SubTables.Count} subtabeller inkluderade.");
                return newLookup;
            }

            return null;
        }
    }
}