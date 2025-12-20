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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
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
                new ExtensionSubstHandler(),
                new ChainingContextualHandler()
            };
            _handlers = handlers.ToDictionary(h => h.LookupType);
        }

        internal void DiscoverLookup(FontSubsettingContext context, LookupTable lookup)
        {
            if (_handlers.TryGetValue(lookup.LookupType, out var handler))
            {
                handler.Discover(context, lookup, this);
            }
        }

        public void Discover(FontSubsettingContext context)
        {
            context.GsubProcessor = this;
            var gsub = context.OriginalFont.GsubTable;
            if (gsub == null || gsub.LookupList == null) return;

            int previousGlyphCount;
            //context.IncludedGlyphs.Add(447);
            //context.IncludedGlyphs.Add(445);
            //context.IncludedGlyphs.Add(449);
            do
            {
                previousGlyphCount = context.IncludedGlyphs.Count;

                foreach (var lookup in gsub.LookupList.Lookups)
                {
                    // Använd den nya centrala metoden
                    DiscoverLookup(context, lookup);
                }

            } while (context.IncludedGlyphs.Count > previousGlyphCount);
        }

        public void Rewrite(FontSubsettingContext context)
        {
            context.GsubProcessor = this;
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

        internal IGsubLookupHandler GetHandler(ushort lookupType)
        {
            if (_handlers.TryGetValue(lookupType, out var handler))
            {
                return handler;
            }
            return null;
        }

        public LookupTable RewriteLookup(FontSubsettingContext context, LookupTable lookup)
        {
            // Hitta rätt handler (SingleSubst, ChainingContextual, etc.)
            if (_handlers.TryGetValue(lookup.LookupType, out var handler))
            {
                // Kör den specifika Rewrite-logiken som vi har jobbat på
                return handler.Rewrite(context, lookup);
            }

            // Om vi inte har en handler för denna typ (t.ex. Type 2 eller 3), 
            // returnerar vi null så att LookupListTable kan skapa en tom platshållare
            return null;
        }
    }
}