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
using EPPlus.Fonts.OpenType.Tables.Gsub.Handlers;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class GsubSubsetProcessor : IFontSubsetProcessor
    {
        private readonly Dictionary<ushort, IGsubLookupHandler> _handlers;

        public GsubSubsetProcessor()
        {
            var handlers = new IGsubLookupHandler[]
            {
                new SingleSubstHandler(),
                new LigatureSubstHandler(),
            };
            _handlers = handlers.ToDictionary(h => h.LookupType);
        }

        public void Discover(FontSubsettingContext context)
        {
            var gsub = context.OriginalFont.GsubTable;
            if (gsub == null || gsub.LookupList == null) return;

            int previousGlyphCount;
            do
            {
                previousGlyphCount = context.IncludedGlyphs.Count;

                foreach (var lookup in gsub.LookupList.Lookups)
                {
                    if (_handlers.TryGetValue(lookup.LookupType, out var handler))
                    {
                        handler.Discover(context, lookup);
                    }
                }

                // Continue as long as new glyphs are found (e.g., a ligature that triggers another substitution rule)
            } while (context.IncludedGlyphs.Count > previousGlyphCount);
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var oldGsub = context.OriginalFont.GsubTable;
            if (oldGsub == null) return;

            // Här aktiverar vi din nya kedja!
            // Istället för att bygga allt manuellt här, låter vi GsubTable sköta det.
            var newGsub = oldGsub.Rewrite(context);

            if (newGsub != null)
            {
                context.SubsetFont.AddOrReplaceTable(newGsub);
            }
        }
    }
}