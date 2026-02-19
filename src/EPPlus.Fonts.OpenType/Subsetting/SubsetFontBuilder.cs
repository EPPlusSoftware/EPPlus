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
using EPPlus.Fonts.OpenType.Subsetting.Processors;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Orchestrates the process of creating a subset of an OpenType font.
    /// </summary>
    internal class SubsetFontBuilder
    {
        // The order of processors is critical: Discovery must precede Extraction and Rewriting.
        private static IEnumerable<IFontSubsetProcessor> Processors => new List<IFontSubsetProcessor>
        {
            // PHASE 1: DISCOVERY - Identify all required Glyph IDs (GIDs)
            new CmapSubsetProcessor(),      // Maps Unicode characters to initial GIDs
            new GsubSubsetProcessor(),      // Identifies additional GIDs needed for substitutions/ligatures
            new GposSubsetProcessor(),      // ← Processes glyph positioning (kerning, accents, etc.)

            // PHASE 2: DATA EXTRACTION - Retrieve glyph outlines and metrics
            new GlyfAndLocaSubsetProcessor(), 

            // PHASE 3: METADATA & TABLES - Update remaining font tables
            new MaxpSubsetProcessor(),
            new HeadSubsetProcessor(),
            new NameSubsetProcessor(),
            new HheaSubsetProcessor(),
            new HmtxSubsetProcessor(),
            new VheaSubsetProcessor(),
            new VmtxSubsetProcessor(),
            new Os2SubsetProcessor(),
            new PostSubsetProcessor(),
            new KernSubsetProcessor()
        };

        /// <summary>
        /// Creates a new <see cref="OpenTypeFont"/> containing only the necessary data for the specified characters.
        /// </summary>
        public OpenTypeFont CreateSubset(OpenTypeFont originalFont, IEnumerable<int> unicodeChars)
        {
            var context = new FontSubsettingContext(originalFont, unicodeChars);
            var processors = Processors;

            // Step 1: Discovery Phase - All processors identify required glyphs
            foreach (var processor in processors)
            {
                processor.Discover(context);
            }

            // Step 2: Build Glyph ID Mapping
            BuildGlyphMapping(context);

            // Step 3: Rewrite Phase - Reconstruct tables using the new Glyph IDs
            foreach (var processor in processors)
            {
                processor.Rewrite(context);
            }

            context.SubsetFont.UsedCodePointsForSubset = new List<uint>(context.UsedCodePoints);

            context.SubsetFont.SubsetGlyphMapping = new Dictionary<ushort, ushort>(context.OldToNewGlyphId);

            return context.SubsetFont;
        }

        // Creates a deterministic mapping between old and new Glyph IDs.
        private void BuildGlyphMapping(FontSubsettingContext context)
        {
            var sortedGlyphs = new List<ushort>(context.IncludedGlyphs);
            sortedGlyphs.Sort();

            for (ushort newId = 0; newId < sortedGlyphs.Count; newId++)
            {
                ushort oldId = sortedGlyphs[newId];
                context.OldToNewGlyphId[oldId] = newId;
                context.NewToOldGlyphId.Add(oldId);
            }
        }
    }
}