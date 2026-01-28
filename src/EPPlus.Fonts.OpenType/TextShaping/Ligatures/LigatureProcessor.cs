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
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.TextShaping.Ligatures
{
    internal class LigatureProcessor
    {
        private readonly List<LookupTable> _ligaLookups;

        public LigatureProcessor(OpenTypeFont font)
        {
            _font = font;
            if (font.GsubTable != null)
            {
                _ligaLookups = FindLookupsForFeature(font.GsubTable, "liga");
            }
            else
            {
                _ligaLookups = new List<LookupTable>();
            }
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
                ApplyLigaturesInPlace(glyphs);
            }

            return glyphs;
        }

        internal void ApplyLigaturesInPlace(List<ShapedGlyph> glyphs)
        {
            if (_ligaLookups.Count == 0) return;

            foreach (var lookup in _ligaLookups)
            {
                if (lookup.LookupType != 4) continue;

                int i = 0;
                while (i < glyphs.Count)
                {
                    bool substituted = false;

                    foreach (var subtableObj in lookup.SubTables)
                    {
                        if (subtableObj is not LigatureSubstSubTable subtable) continue;

                        if (TryApplyLigatureInPlace(glyphs, i, subtable, out int consumed))
                        {
                            substituted = true;
                            i += consumed; // Oftast 1 efter ersättning
                            break;         // Första match vinner – hoppa ur
                        }
                    }

                    if (!substituted) i++;
                }
            }
        }

        private bool TryApplyLigatureInPlace(
            List<ShapedGlyph> glyphs,
            int startIndex,
            LigatureSubstSubTable subtable,
            out int componentsConsumed)
        {
            componentsConsumed = 0;

            if (startIndex >= glyphs.Count) return false;

            ushort first = glyphs[startIndex].GlyphId;
            int covIdx = subtable.Coverage.GetGlyphIndex(first);
            if (covIdx < 0) return false;

            if (!subtable.LigatureSets.TryGetValue(first, out var ligSet) || ligSet?.Ligatures.Count == 0)
                return false;

            // Försök längre ligaturer först (rekommenderas av OpenType-spec)
            var sortedLigs = ligSet.Ligatures
                .OrderByDescending(l => 1 + (l.Components?.Length ?? 0))
                .ToList();

            foreach (var lig in sortedLigs)
            {
                int compCount = 1 + (lig.Components?.Length ?? 0);
                if (startIndex + compCount > glyphs.Count) continue;

                bool match = true;
                for (int j = 0; j < lig.Components?.Length; j++)
                {
                    if (glyphs[startIndex + 1 + j].GlyphId != lig.Components[j])
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                {
                    var ligGlyph = CreateLigatureGlyph(glyphs, startIndex, (byte)compCount, lig.LigatureGlyph);

                    // MUTERA DIREKT
                    glyphs.RemoveRange(startIndex, compCount);
                    glyphs.Insert(startIndex, ligGlyph);

                    componentsConsumed = 1; // ligatur tar platsen → nästa steg flyttar förbi den
                    return true;
                }
            }

            return false;
        }


        /// <summary>
        /// Finds all lookups associated with a feature tag.
        /// </summary>
        private List<LookupTable> FindLookupsForFeature(GsubTable gsub, string featureTag)
        {
            var lookups = new List<LookupTable>();

            if (gsub?.FeatureList?.FeatureRecords == null)
                return lookups;

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
        /// Creates a new shaped glyph for a ligature, combining metrics from components.
        /// </summary>
        private ShapedGlyph CreateLigatureGlyph(
            List<ShapedGlyph> glyphs,
            int startIndex,
            byte componentCount,
            ushort ligatureGlyphId)
        {
            // Get advance width for ligature glyph
            var advanceWidth = (short)_font.HmtxTable.GetAdvanceWidth(ligatureGlyphId);

            // Preserve cluster index from first component
            var clusterIndex = glyphs[startIndex].ClusterIndex;

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
