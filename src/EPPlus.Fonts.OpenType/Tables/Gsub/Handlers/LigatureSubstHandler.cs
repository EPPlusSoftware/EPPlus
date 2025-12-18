/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System.Collections.Generic;
using System.Linq;

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
