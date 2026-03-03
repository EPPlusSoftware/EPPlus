/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS subset processor
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using EPPlus.Fonts.OpenType.Tables.Gpos.Handlers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Processor for subsetting GPOS (Glyph Positioning) table.
    /// Implements the two-phase subsetting pattern: Discover → Rewrite.
    /// </summary>
    internal class GposSubsetProcessor : IFontSubsetProcessor
    {
        private readonly Dictionary<ushort, IGposLookupHandler> _handlers;

        public GposSubsetProcessor()
        {
            var handlers = new IGposLookupHandler[]
            {
                new SinglePosHandler(),      // Type 1
                new PairPosHandler(),        // Type 2
                new MarkToBaseHandler(),     // Type 4
                new ExtensionPosHandler()    // Type 9 
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

        /// <summary>
        /// Phase 1: Discover dependencies.
        /// For GPOS, there are typically no additional glyphs to discover
        /// (unlike GSUB where ligatures require component glyphs).
        /// GPOS only positions existing glyphs.
        /// </summary>
        public void Discover(FontSubsettingContext context)
        {
            context.GposProcessor = this;
            var gpos = context.OriginalFont.GposTable;
            if (gpos == null) return;


            // GPOS typically doesn't discover new glyphs
            // But we still iterate through lookups in case future handlers need it
            foreach (var lookup in gpos.LookupList.Lookups)
            {
                DiscoverLookup(context, lookup);
            }
        }

        /// <summary>
        /// Phase 2: Rewrite the GPOS table with subsetted lookups and remapped glyph IDs.
        /// </summary>
        public void Rewrite(FontSubsettingContext context)
        {
            context.GposProcessor = this;
            var oldGpos = context.OriginalFont.GposTable;
            if (oldGpos == null) return;

            // Count Type 4 (MarkToBase) lookups
            int markToBaseCount = 0;
            foreach (var lookup in oldGpos.LookupList.Lookups)
            {
                if (lookup.LookupType == 4 ||
                    (lookup.LookupType == 9 && lookup.SubTables.Count > 0 &&
                     lookup.SubTables[0] is MarkToBaseSubTableFormat1))
                {
                    markToBaseCount++;
                }
            }

            var newGpos = oldGpos.Rewrite(context);

            if (newGpos != null)
            {
                context.SubsetFont.AddOrReplaceTable(newGpos);
            }
        }

        internal IGposLookupHandler GetHandler(ushort lookupType)
        {
            if (_handlers.TryGetValue(lookupType, out var handler))
            {
                return handler;
            }
            return null;
        }

        public LookupTable RewriteLookup(FontSubsettingContext context, LookupTable lookup)
        {
            // Find the right handler (SinglePos, PairPos, MarkToBase, etc.)
            if (_handlers.TryGetValue(lookup.LookupType, out var handler))
            {
                // Run the specific Rewrite logic
                return handler.Rewrite(context, lookup);
            }

            // If we don't have a handler for this type (e.g., Type 3, 5, 6),
            // return null so LookupListTable can create an empty placeholder
            return null;
        }
    }
}