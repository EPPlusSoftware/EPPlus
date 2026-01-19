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
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal class SingleSubstHandler : IGsubLookupHandler
    {
        public ushort LookupType => 1;

        public void Discover(FontSubsettingContext context, LookupTable lookup, GsubSubsetProcessor processor)
        {
            // Loop until no new glyphs are added to capture transitive substitution chains
            bool glyphsAdded;
            do
            {
                glyphsAdded = false;
                var currentGlyphs = context.IncludedGlyphs.ToArray();

                foreach (var subtable in lookup.SubTables.OfType<SingleSubstSubTable>())
                {
                    foreach (ushort gid in currentGlyphs)
                    {
                        ushort substitute = subtable.GetSubstitution(gid);

                        if (substitute != 0 && !context.IncludedGlyphs.Contains(substitute))
                        {
                            context.IncludedGlyphs.Add(substitute);
                            glyphsAdded = true;
                        }
                    }
                }
            } while (glyphsAdded);
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newLookup = new LookupTable { LookupType = 1, LookupFlag = oldLookup.LookupFlag, SubTables = new List<FontTableElement>() };

            foreach (var subtable in oldLookup.SubTables.OfType<SingleSubstSubTable>())
            {
                var oldMappings = GetMappings(subtable);
                var validMappings = new List<GsubRewriteEntry>();

                // Remap to subset glyph IDs, keeping only mappings where both input and output glyphs are included
                foreach (var map in oldMappings)
                {
                    if (context.OldToNewGlyphId.TryGetValue(map.Key, out ushort newInputGid) &&
                        context.OldToNewGlyphId.TryGetValue(map.Value, out ushort newOutputGid))
                    {
                        validMappings.Add(new GsubRewriteEntry { NewInput = newInputGid, NewOutput = newOutputGid });
                    }
                }

                if (validMappings.Count > 0)
                {
                    validMappings.Sort((a, b) => a.NewInput.CompareTo(b.NewInput));
                    var newSub = new SingleSubstSubTableFormat2
                    {
                        SubstituteGlyphIDs = validMappings.Select(m => m.NewOutput).ToArray(),
                        Coverage = CoverageTableFormat2.CreateCoverageFormat2(validMappings.Select(m => m.NewInput).ToList()),
                        GlyphCount = (ushort)validMappings.Count,
                        SubtableFormat = 2
                    };
                    newLookup.SubTables.Add(newSub);
                }
            }
            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }

        private Dictionary<ushort, ushort> GetMappings(SingleSubstSubTable subtable)
        {
            var mappings = new Dictionary<ushort, ushort>();
            var inputGids = subtable.Coverage.GetCoveredGlyphs();

            foreach (var oldGid in inputGids)
            {
                ushort substitutedGid = subtable.GetSubstitution(oldGid);

                if (substitutedGid != 0)
                {
                    mappings[oldGid] = substitutedGid;
                }
            }

            return mappings;
        }
    }
}