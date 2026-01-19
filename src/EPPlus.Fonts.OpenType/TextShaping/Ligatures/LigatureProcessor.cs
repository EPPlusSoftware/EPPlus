/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/15/2025         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Ligatures
{
    internal class LigatureProcessor
    {
        public LigatureProcessor(OpenTypeFont font)
        {
            _font = font;
        }

        private readonly OpenTypeFont _font;

        /// <summary>
        /// Applies standard ligature substitutions (fi, ff, ffi, ffl, etc.).
        /// Processes glyphs left-to-right, replacing sequences with ligature glyphs.
        /// </summary>
        internal List<ShapedGlyph> ApplyLigatures(List<ShapedGlyph> glyphs)
        {
            var gsub = _font.GsubTable;
            if (gsub == null)
                return glyphs;

            // Find "liga" feature
            var ligaLookups = FindLookupsForFeature(gsub, "liga");
            if (ligaLookups.Count == 0)
                return glyphs;

            // Apply each lookup in order
            foreach (var lookup in ligaLookups)
            {
                glyphs = ApplyLigatureLookup(glyphs, lookup);
            }

            return glyphs;
        }

        /// <summary>
        /// Finds all lookups associated with a feature tag.
        /// </summary>
        private List<LookupTable> FindLookupsForFeature(GsubTable gsub, string featureTag)
        {
            var lookups = new List<LookupTable>();

            foreach (var featureRecord in gsub.FeatureList.FeatureRecords)
            {
                if (featureRecord.FeatureTag.Value == featureTag)
                {
                    var feature = featureRecord.FeatureTable;

                    foreach (var lookupIndex in feature.LookupListIndices)
                    {
                        if (lookupIndex < gsub.LookupList.Lookups.Count)
                        {
                            lookups.Add(gsub.LookupList.Lookups[lookupIndex]);
                        }
                    }
                }
            }

            return lookups;
        }

        /// <summary>
        /// Applies a single ligature lookup to the glyph sequence.
        /// Processes left-to-right, replacing matching sequences with ligatures.
        /// </summary>
        private List<ShapedGlyph> ApplyLigatureLookup(List<ShapedGlyph> glyphs, LookupTable lookup)
        {
            if (lookup.LookupType != 4) // Must be Ligature Substitution
                return glyphs;

            var result = new List<ShapedGlyph>();
            int i = 0;

            while (i < glyphs.Count)
            {
                bool substituted = false;

                // Try each subtable
                foreach (var subtable in lookup.SubTables)
                {
                    if (subtable is LigatureSubstSubTable ligSubtable)
                    {
                        // Try to match ligature starting at position i
                        if (TryApplyLigature(glyphs, i, ligSubtable, out var ligatureGlyph, out int componentsConsumed))
                        {
                            result.Add(ligatureGlyph);
                            i += componentsConsumed;
                            substituted = true;
                            break; // Found a match, move to next position
                        }
                    }
                }

                if (!substituted)
                {
                    // No ligature found, keep original glyph
                    result.Add(glyphs[i]);
                    i++;
                }
            }

            return result;
        }

        /// <summary>
        /// Attempts to find and apply a ligature substitution starting at the given position.
        /// </summary>
        private bool TryApplyLigature(
             List<ShapedGlyph> glyphs,
             int startIndex,
             LigatureSubstSubTable subtable,
             out ShapedGlyph ligatureGlyph,
             out int componentsConsumed)
        {
            ligatureGlyph = null;
            componentsConsumed = 0;

            if (startIndex >= glyphs.Count)
                return false;

            ushort firstGlyph = glyphs[startIndex].GlyphId;

            // Check if first glyph is in coverage
            int coverageIndex = subtable.Coverage.GetGlyphIndex(firstGlyph);
            if (coverageIndex < 0)
                return false;

            // LigatureSets is a Dictionary<ushort, LigatureSet>
            // Key is the GLYPH ID, not coverage index!
            if (!subtable.LigatureSets.TryGetValue(firstGlyph, out var ligatureSet))
                return false;

            if (ligatureSet?.Ligatures == null)
                return false;

            // Try each ligature in the set
            foreach (var ligature in ligatureSet.Ligatures)
            {
                int componentCount = 1 + (ligature.Components?.Length ?? 0);

                // Check if we have enough glyphs remaining
                if (startIndex + componentCount > glyphs.Count)
                    continue;

                // Check if all component glyphs match
                bool matches = true;

                if (ligature.Components != null)
                {
                    for (int j = 0; j < ligature.Components.Length; j++)
                    {
                        if (glyphs[startIndex + 1 + j].GlyphId != ligature.Components[j])
                        {
                            matches = false;
                            break;
                        }
                    }
                }

                if (matches)
                {
                    // Found a match! Create ligature glyph
                    ligatureGlyph = CreateLigatureGlyph(glyphs, startIndex, componentCount, ligature.LigatureGlyph);
                    componentsConsumed = componentCount;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Creates a new shaped glyph for a ligature, combining metrics from components.
        /// </summary>
        private ShapedGlyph CreateLigatureGlyph(
            List<ShapedGlyph> glyphs,
            int startIndex,
            int componentCount,
            ushort ligatureGlyphId)
        {
            // Get advance width for ligature glyph
            int advanceWidth = _font.HmtxTable.GetAdvanceWidth(ligatureGlyphId);

            // Preserve cluster index from first component
            int clusterIndex = glyphs[startIndex].ClusterIndex;

            return new ShapedGlyph
            {
                GlyphId = ligatureGlyphId,
                XAdvance = advanceWidth,
                YAdvance = 0,
                XOffset = 0,
                YOffset = 0,
                ClusterIndex = clusterIndex,
                CharCount = componentCount // Ligature represents multiple characters
            };
        }
    }
}
