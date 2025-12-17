using EPPlus.Fonts.OpenType.Subsetting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal class LigatureSubstHandler : IGsubLookupHandler
    {
        public ushort LookupType => 4;

        public void Discover(FontSubsettingContext context, LookupTable lookup)
        {
            foreach (var subtable in lookup.SubTables.OfType<LigatureSubstSubTable>())
            {
                if (subtable.LigatureSets == null) continue;

                foreach (var kvp in subtable.LigatureSets)
                {
                    ushort firstGlyph = kvp.Key;
                    if (!context.IncludedGlyphs.Contains(firstGlyph)) continue;

                    foreach (var lig in kvp.Value.Ligatures)
                    {
                        bool allComponentsPresent = true;
                        if (lig.Components != null)
                        {
                            foreach (ushort compGid in lig.Components)
                            {
                                if (!context.IncludedGlyphs.Contains(compGid))
                                {
                                    allComponentsPresent = false;
                                    break;
                                }
                            }
                        }

                        if (allComponentsPresent && !context.IncludedGlyphs.Contains(lig.LigatureGlyph))
                        {
                            context.IncludedGlyphs.Add(lig.LigatureGlyph);
                        }
                    }
                }
            }
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newLookup = new LookupTable { LookupType = 4, LookupFlag = oldLookup.LookupFlag, SubTables = new List<FontTableElement>() };
            foreach (var subtable in oldLookup.SubTables.OfType<LigatureSubstSubTable>())
            {
                var rewritten = subtable.Rewrite(context);
                if (rewritten != null) newLookup.SubTables.Add(rewritten);
            }
            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }
    }
}
