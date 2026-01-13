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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal class LigatureSubstHandler : IGsubLookupHandler
    {
        public ushort LookupType => 4;

        public void Discover(FontSubsettingContext context, LookupTable lookup, GsubSubsetProcessor processor)
        {
            bool addedAny;
            int iterationCount = 0;

            do
            {
                iterationCount++;
                addedAny = false;

                foreach (var subtable in lookup.SubTables)
                {
                    LigatureSubstSubTable ligSubtable = subtable as LigatureSubstSubTable;
                    if (ligSubtable == null || ligSubtable.LigatureSets == null) continue;

                    foreach (var kvp in ligSubtable.LigatureSets)
                    {
                        ushort firstGlyphId = kvp.Key;

                        if (!context.IncludedGlyphs.Contains(firstGlyphId))
                            continue;

                        foreach (var lig in kvp.Value.Ligatures)
                        {
                            if (context.IncludedGlyphs.Contains(lig.LigatureGlyph))
                                continue;

                            // ✅ SIMPLIFIED: Only check base character components (< 400)
                            bool allBaseComponentsExist = true;

                            if (lig.Components != null && lig.Components.Length > 0)
                            {
                                foreach (ushort compGid in lig.Components)
                                {
                                    // ✅ Only validate base characters (< 400)
                                    if (compGid < 400)
                                    {
                                        if (!context.IncludedGlyphs.Contains(compGid))
                                        {
                                            allBaseComponentsExist = false;
                                            break;
                                        }
                                    }
                                    // Ligature components (>= 400) are completely ignored
                                }
                            }

                            if (allBaseComponentsExist)
                            {
                                context.IncludedGlyphs.Add(lig.LigatureGlyph);
                                addedAny = true;
                            }
                        }
                    }
                }
            } while (addedAny);
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {

            var newLookup = new LookupTable
            {
                LookupType = 4,
                LookupFlag = oldLookup.LookupFlag,
                SubTables = new List<FontTableElement>()
            };

            foreach (var oldSubtable in oldLookup.SubTables)
            {
                LigatureSubstSubTable oldLigSubtable = oldSubtable as LigatureSubstSubTable;
                if (oldLigSubtable == null) continue;
                if (oldLigSubtable.LigatureSets == null) continue;

                var newSubTable = new LigatureSubstSubTable();
                newSubTable.SubtableFormat = 1;
                newSubTable.LigatureSets = new Dictionary<ushort, LigatureSetTable>();

                foreach (var kvp in oldLigSubtable.LigatureSets)
                {
                    ushort oldFirstGid = kvp.Key;

                    if (!context.OldToNewGlyphId.TryGetValue(oldFirstGid, out ushort newFirstGid))
                        continue;

                    var rewrittenSet = kvp.Value.Rewrite(context);
                    if (rewrittenSet != null && rewrittenSet.Ligatures.Count > 0)
                    {
                        newSubTable.LigatureSets[newFirstGid] = rewrittenSet;
                    }
                }

                if (newSubTable.LigatureSets.Count > 0)
                {
                    var coveredGids = new List<ushort>(newSubTable.LigatureSets.Keys);
                    coveredGids.Sort();

                    newSubTable.Coverage = new CoverageTableFormat1
                    {
                        CoverageFormat = 1,
                        GlyphArray = coveredGids.ToArray(),
                        GlyphCount = (ushort)coveredGids.Count
                    };

                    newLookup.SubTables.Add(newSubTable);
                }
            }

            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }
    }
}
