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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System;
using System.Collections.Generic;
using System.Linq;

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
                System.Diagnostics.Debug.WriteLine(string.Format("Iteration {0}, current glyph count: {1}", iterationCount, context.IncludedGlyphs.Count));

                addedAny = false;

                foreach (var subtable in lookup.SubTables)
                {
                    LigatureSubstSubTable ligSubtable = subtable as LigatureSubstSubTable;
                    if (ligSubtable == null) continue;
                    if (ligSubtable.LigatureSets == null) continue;

                    foreach (var kvp in ligSubtable.LigatureSets)
                    {
                        ushort firstGlyphId = kvp.Key;

                        if (!context.IncludedGlyphs.Contains(firstGlyphId))
                            continue;

                        foreach (var lig in kvp.Value.Ligatures)
                        {
                            if (context.IncludedGlyphs.Contains(lig.LigatureGlyph))
                                continue;

                            // ✅ FIX: Lägg till ligatur-glyphen OM första tecknet finns
                            // Vi bryr oss INTE om komponenterna just nu - vi lägger till den ändå
                            // eftersom vi behöver den för andra ligaturer
                            context.IncludedGlyphs.Add(lig.LigatureGlyph);
                            addedAny = true;

                            System.Diagnostics.Debug.WriteLine(string.Format("Added ligature GID {0} (first: {1})",
                                lig.LigatureGlyph, firstGlyphId));
                        }
                    }
                }
            } while (addedAny);
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            System.Diagnostics.Debug.WriteLine("=== LigatureSubstHandler.Rewrite START ===");
            System.Diagnostics.Debug.WriteLine(string.Format("Old lookup has {0} subtables", oldLookup.SubTables.Count));

            var newLookup = new LookupTable
            {
                LookupType = 4,
                LookupFlag = oldLookup.LookupFlag,
                SubTables = new List<FontTableElement>()
            };

            foreach (var subtable in oldLookup.SubTables)
            {
                LigatureSubstSubTable ligSubtable = subtable as LigatureSubstSubTable;
                if (ligSubtable == null || ligSubtable.LigatureSets == null) continue;

                System.Diagnostics.Debug.WriteLine(string.Format("Processing subtable with {0} LigatureSets", ligSubtable.LigatureSets.Count));

                var newSubTable = new LigatureSubstSubTable();
                newSubTable.SubtableFormat = 1;  // ✅ LÄGG TILL DENNA RAD
                newSubTable.LigatureSets = new Dictionary<ushort, LigatureSetTable>();

                foreach (var kvp in ligSubtable.LigatureSets)
                {
                    ushort oldFirstGlyphId = kvp.Key;

                    System.Diagnostics.Debug.WriteLine(string.Format("  Checking old GID {0}, in mapping: {1}",
                        oldFirstGlyphId,
                        context.OldToNewGlyphId.ContainsKey(oldFirstGlyphId)));

                    if (context.OldToNewGlyphId.TryGetValue(oldFirstGlyphId, out ushort newFirstGlyphId))
                    {
                        var rewrittenSet = kvp.Value.Rewrite(context);

                        if (rewrittenSet != null && rewrittenSet.Ligatures.Count > 0)
                        {
                            newSubTable.LigatureSets[newFirstGlyphId] = rewrittenSet;
                            System.Diagnostics.Debug.WriteLine(string.Format("    ✅ Added LigatureSet for new GID {0} with {1} ligatures",
                                newFirstGlyphId,
                                rewrittenSet.Ligatures.Count));
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine(string.Format("Subtable result: {0} LigatureSets", newSubTable.LigatureSets.Count));

                if (newSubTable.LigatureSets.Count > 0)
                {
                    // ✅ SKAPA COVERAGE
                    var coveredGids = new List<ushort>(newSubTable.LigatureSets.Keys);
                    coveredGids.Sort();

                    newSubTable.Coverage = new CoverageTableFormat1
                    {
                        CoverageFormat = 1,
                        GlyphArray = coveredGids.ToArray(),
                        GlyphCount = (ushort)coveredGids.Count
                    };

                    newLookup.SubTables.Add(newSubTable);
                    System.Diagnostics.Debug.WriteLine(string.Format("    ✅ Added subtable with Coverage for {0} glyphs", coveredGids.Count));
                }
            }

            System.Diagnostics.Debug.WriteLine(string.Format("=== LigatureSubstHandler.Rewrite END: {0} subtables ===", newLookup.SubTables.Count));
            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }
    }
}
