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
  09/07/2026         EPPlus Software AB           Filter by ScriptList/LangSys, not just tag
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.TextShaping.Ligatures
{
    internal class LigatureProcessor
    {
        private readonly struct IndexedLookup
        {
            public readonly int FeatureIndex;
            public readonly LookupTable Lookup;

            public IndexedLookup(int featureIndex, LookupTable lookup)
            {
                FeatureIndex = featureIndex;
                Lookup = lookup;
            }
        }

        private readonly OpenTypeFont _font;
        private readonly List<IndexedLookup> _ligaLookups;
        private readonly Dictionary<string, HashSet<int>> _activeIndexCache = new Dictionary<string, HashSet<int>>();

        public LigatureProcessor(OpenTypeFont font)
        {
            _font = font;
            if (font.GsubTable != null)
            {
                _ligaLookups = FindLookupsForFeature(font.GsubTable, "liga");
            }
            else
            {
                _ligaLookups = new List<IndexedLookup>();
            }
        }

        /// <summary>
        /// Applies standard ligature substitutions (fi, ff, ffi, ffl, etc.).
        /// Processes glyphs left-to-right, replacing sequences with ligature glyphs.
        /// </summary>
        internal List<ShapedGlyph> ApplyLigatures(List<ShapedGlyph> glyphs, string script, string language)
        {
            var gsub = _font.GsubTable;
            if (gsub == null)
                return glyphs;

            if (_ligaLookups.Count == 0)
                return glyphs;

            ApplyLigaturesInPlace(glyphs, script, language);

            return glyphs;
        }

        internal void ApplyLigaturesInPlace(List<ShapedGlyph> glyphs, string script, string language)
        {
            if (_ligaLookups.Count == 0) return;

            HashSet<int> activeIndices = GetActiveIndices(script, language);

            foreach (var entry in _ligaLookups)
            {
                // null activeIndices means "no ScriptList to filter by" - keep every entry,
                // matching the previous behavior rather than discarding ligatures we cannot resolve.
                if (activeIndices != null && !activeIndices.Contains(entry.FeatureIndex))
                {
                    continue;
                }

                var lookup = entry.Lookup;
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
                            i += consumed; // Usually 1 after a substitution
                            break;         // First match wins - break out
                        }
                    }

                    if (!substituted) i++;
                }
            }
        }

        private HashSet<int> GetActiveIndices(string script, string language)
        {
            string cacheKey = (script ?? string.Empty) + "|" + (language ?? string.Empty);

            if (!_activeIndexCache.TryGetValue(cacheKey, out var indices))
            {
                indices = ScriptFeatureResolver.GetActiveFeatureIndices(_font.GsubTable?.ScriptList, script, language);
                _activeIndexCache[cacheKey] = indices;
            }

            return indices;
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

            // Try longer ligatures first, as recommended by the OpenType spec
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

                    // MUTATE IN PLACE
                    glyphs.RemoveRange(startIndex, compCount);
                    glyphs.Insert(startIndex, ligGlyph);

                    componentsConsumed = 1; // the ligature takes its place - the next step moves past it
                    return true;
                }
            }

            return false;
        }


        /// <summary>
        /// Finds all lookups associated with a feature tag, together with each one's original
        /// FeatureList index, across all scripts. Two FeatureRecords can legitimately share a
        /// tag (one per script); both are kept so ApplyLigaturesInPlace can filter by script at
        /// call time instead of the constructor baking in whichever script happened to be active
        /// when this processor was built.
        /// </summary>
        private List<IndexedLookup> FindLookupsForFeature(GsubTable gsub, string featureTag)
        {
            var lookups = new List<IndexedLookup>();

            if (gsub?.FeatureList?.FeatureRecords == null)
                return lookups;

            var featureRecords = gsub.FeatureList.FeatureRecords;

            for (int featureIndex = 0; featureIndex < featureRecords.Count; featureIndex++)
            {
                var featureRecord = featureRecords[featureIndex];

                if (featureRecord.FeatureTag.Value != featureTag)
                    continue;

                var feature = featureRecord.FeatureTable;

                foreach (var lookupIndex in feature.LookupListIndices)
                {
                    if (lookupIndex < gsub.LookupList.Lookups.Count)
                    {
                        lookups.Add(new IndexedLookup(featureIndex, gsub.LookupList.Lookups[lookupIndex]));
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
            var baseAdvance = (short)_font.HmtxTable.GetAdvanceWidth(ligatureGlyphId);
            var clusterIndex = glyphs[startIndex].ClusterIndex;

            return new ShapedGlyph
            {
                GlyphId = ligatureGlyphId,
                BaseAdvance = baseAdvance,      // Base advance for ligature
                XAdvance = baseAdvance,         // Will be adjusted by positioning
                YAdvance = 0,
                XOffset = 0,
                YOffset = 0,
                ClusterIndex = clusterIndex,
                CharCount = componentCount
            };
        }
    }
}