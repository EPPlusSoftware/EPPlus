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
    internal class SingleSubstHandler : IGsubLookupHandler
    {
        public ushort LookupType => 1;

        public void Discover(FontSubsettingContext context, LookupTable lookup)
        {
            // We iterate using a copy of the current glyphs to check if any existing glyph 
            // triggers a substitution. The GsubSubsetProcessor manages the iterative state
            // by monitoring if context.IncludedGlyphs grows.
            var currentGlyphs = context.IncludedGlyphs.ToArray();
            foreach (var subtable in lookup.SubTables.OfType<SingleSubstSubTable>())
            {
                foreach (ushort gid in currentGlyphs)
                {
                    ushort substitute = subtable.GetSubstitution(gid);
                    if (substitute != 0 && !context.IncludedGlyphs.Contains(substitute))
                    {
                        context.IncludedGlyphs.Add(substitute);
                        // The addition to the HashSet will be detected by the 
                        // do-while loop in GsubSubsetProcessor.
                    }
                }
            }
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            var newLookup = new LookupTable
            {
                LookupType = 1,
                LookupFlag = oldLookup.LookupFlag,
                SubTables = new List<FontTableElement>()
            };

            foreach (var subtable in oldLookup.SubTables.OfType<SingleSubstSubTable>())
            {
                var rewritten = subtable.Rewrite(context);
                if (rewritten != null)
                {
                    newLookup.SubTables.Add(rewritten);
                }
            }

            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }
    }
}