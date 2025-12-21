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
using EPPlus.Fonts.OpenType.Tables.Common.Coverage;
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
            // Tips: Vi kan loopa tills IncludedGlyphs slutar växa för att fånga kedjereaktioner
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

                        // Om vi hittar en giltig ersättare som vi inte har än
                        if (substitute != 0 && !context.IncludedGlyphs.Contains(substitute))
                        {
                            context.IncludedGlyphs.Add(substitute);
                            glyphsAdded = true;
                        }
                    }
                }
            } while (glyphsAdded); // Kör igen om vi hittade nya glyfer (viktigt för kedjade byten)
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newLookup = new LookupTable { LookupType = 1, LookupFlag = oldLookup.LookupFlag, SubTables = new List<FontTableElement>() };

            foreach (var subtable in oldLookup.SubTables.OfType<SingleSubstSubTable>())
            {
                // 1. Extrahera mappningarna med hjälpmetoden ovan
                var oldMappings = GetMappings(subtable);
                var validMappings = new List<GsubRewriteEntry>();

                // 2. Mappa om dem till subset-GIDs
                foreach (var map in oldMappings)
                {
                    if (context.OldToNewGlyphId.TryGetValue(map.Key, out ushort newInputGid) &&
                        context.OldToNewGlyphId.TryGetValue(map.Value, out ushort newOutputGid))
                    {
                        validMappings.Add(new GsubRewriteEntry { NewInput = newInputGid, NewOutput = newOutputGid });
                    }
                }

                // 3. Skapa den nya subtabellen om vi har kvar några mappningar
                if (validMappings.Count > 0)
                {
                    validMappings.Sort((a, b) => a.NewInput.CompareTo(b.NewInput));
                    var newSub = new SingleSubstSubTableFormat2
                    {
                        SubstituteGlyphIDs = validMappings.Select(m => m.NewOutput).ToArray(),
                        Coverage = CoverageTableFormat2.CreateCoverageFormat2(validMappings.Select(m => m.NewInput).ToList()),
                        GlyphCount = (ushort)validMappings.Count, // <--- VIKTIGT!
                        SubtableFormat = 2
                    };
                    newLookup.SubTables.Add(newSub);
                    System.Diagnostics.Debug.WriteLine($"SingleSubst: Behåller {validMappings.Count} mappningar.");
                }
            }
            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }

        private Dictionary<ushort, ushort> GetMappings(SingleSubstSubTable subtable)
        {
            var mappings = new Dictionary<ushort, ushort>();

            // Hämta alla GID som denna subtable täcker
            var inputGids = subtable.Coverage.GetCoveredGlyphs();

            foreach (var oldGid in inputGids)
            {
                // Använd den existerande abstrakta metoden som sköter 
                // både Delta (Format 1) och Array (Format 2) åt oss!
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