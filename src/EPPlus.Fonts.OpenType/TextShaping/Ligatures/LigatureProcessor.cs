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
  09/07/2026         EPPlus Software AB           Read feature tags from GsubFeatures instead of
                                                   hardcoded "liga"; unwrap extension-wrapped (Type
                                                   7) ligature lookups, which GSUB does not flatten
                                                   the way GPOS does
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

        // Feature tag ("liga", "dlig", "clig", "rlig", ...) -> every ligature-relevant lookup
        // recorded under that tag, across ALL scripts, together with each lookup's original
        // FeatureList index so ApplyLigaturesInPlace can filter by script at call time. Built once
        // for every tag the font defines, not just "liga" - which tags are actually applied is
        // decided per call via the featureTags parameter, driven by ShapingOptions.GsubFeatures.
        private readonly Dictionary<string, List<IndexedLookup>> _lookupsByFeatureTag;
        private readonly Dictionary<string, HashSet<int>> _activeIndexCache = new Dictionary<string, HashSet<int>>();

        public LigatureProcessor(OpenTypeFont font)
        {
            _font = font;
            _lookupsByFeatureTag = font.GsubTable != null
                ? BuildLookupsByFeatureTag(font.GsubTable)
                : new Dictionary<string, List<IndexedLookup>>();
        }

        /// <summary>
        /// Applies ligature substitutions for the given feature tags (e.g. "liga", "dlig").
        /// Processes glyphs left-to-right, replacing sequences with ligature glyphs.
        /// </summary>
        internal List<ShapedGlyph> ApplyLigatures(
            List<ShapedGlyph> glyphs, List<string> featureTags, string script, string language)
        {
            if (_font.GsubTable == null)
                return glyphs;

            ApplyLigaturesInPlace(glyphs, featureTags, script, language);

            return glyphs;
        }

        /// <summary>
        /// Applies ligature substitutions for the given feature tags directly onto <paramref
        /// name="glyphs"/>.
        /// </summary>
        /// <param name="featureTags">
        /// Feature tags to apply, e.g. ["liga", "clig"]. Only tags the font actually defines a
        /// ligature lookup for have any effect - a tag with no matching lookup is silently
        /// skipped, the same way an unmatched glyph sequence is.
        /// </param>
        /// <param name="script">
        /// OpenType script tag (e.g. "latn"). Pass null to fall back to unfiltered lookup, which
        /// reproduces the previous behavior for callers that have no script to give.
        /// </param>
        /// <param name="language">OpenType language-system tag, or null for the script's default.</param>
        internal void ApplyLigaturesInPlace(
            List<ShapedGlyph> glyphs, List<string> featureTags, string script, string language)
        {
            if (featureTags == null || featureTags.Count == 0) return;
            if (_lookupsByFeatureTag.Count == 0) return;

            HashSet<int> activeIndices = GetActiveIndices(script, language);

            foreach (var tag in featureTags)
            {
                if (!_lookupsByFeatureTag.TryGetValue(tag, out var entries))
                {
                    continue;
                }

                foreach (var entry in entries)
                {
                    // null activeIndices means "no ScriptList to filter by" - keep every entry,
                    // matching the previous behavior rather than discarding ligatures we cannot
                    // resolve.
                    if (activeIndices != null && !activeIndices.Contains(entry.FeatureIndex))
                    {
                        continue;
                    }

                    ApplyLookup(glyphs, entry.Lookup);
                }
            }
        }

        private void ApplyLookup(List<ShapedGlyph> glyphs, LookupTable lookup)
        {
            int i = 0;
            while (i < glyphs.Count)
            {
                bool substituted = false;

                foreach (var subtableObj in lookup.SubTables)
                {
                    var subtable = UnwrapLigatureSubtable(subtableObj);
                    if (subtable == null) continue;

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

        /// <summary>
        /// Returns the LigatureSubstSubTable a GSUB subtable represents, unwrapping Extension
        /// Substitution (Type 7) if needed. Unlike GPOS, whose loader flattens extension-wrapped
        /// lookups so LookupType 9 never actually appears on a loaded lookup, GSUB keeps the
        /// wrapper: an extension-wrapped ligature lookup has LookupType 7 with
        /// ExtensionSubstSubTable entries in SubTables, each pointing at the real subtable via
        /// ExtendedSubTable. Returns null for anything that is not, directly or via unwrapping, a
        /// LigatureSubstSubTable.
        /// </summary>
        private static LigatureSubstSubTable UnwrapLigatureSubtable(Tables.FontTableElement subtableObj)
        {
            if (subtableObj is LigatureSubstSubTable direct)
            {
                return direct;
            }

            if (subtableObj is ExtensionSubstSubTable extension
                && extension.ExtensionLookupType == 4
                && extension.ExtendedSubTable is LigatureSubstSubTable wrapped)
            {
                return wrapped;
            }

            return null;
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
        /// Builds a map from feature tag to every ligature-relevant lookup recorded under that
        /// tag (LookupType 4, or 7 wrapping 4), together with each lookup's original FeatureList
        /// index. Scans every FeatureRecord in the font, not just "liga" - which tags matter is
        /// decided later, per call, by ApplyLigaturesInPlace's featureTags parameter. Two
        /// FeatureRecords can legitimately share a tag (one per script); both are kept so the
        /// script filter in ApplyLigaturesInPlace has something to choose between.
        /// </summary>
        private static Dictionary<string, List<IndexedLookup>> BuildLookupsByFeatureTag(GsubTable gsub)
        {
            var map = new Dictionary<string, List<IndexedLookup>>();

            if (gsub?.FeatureList?.FeatureRecords == null || gsub.LookupList == null)
                return map;

            var featureRecords = gsub.FeatureList.FeatureRecords;

            for (int featureIndex = 0; featureIndex < featureRecords.Count; featureIndex++)
            {
                var featureRecord = featureRecords[featureIndex];
                string tag = featureRecord.FeatureTag.Value;
                var feature = featureRecord.FeatureTable;

                if (feature?.LookupListIndices == null)
                    continue;

                foreach (var lookupIndex in feature.LookupListIndices)
                {
                    if (lookupIndex >= gsub.LookupList.Lookups.Count)
                        continue;

                    var lookup = gsub.LookupList.Lookups[lookupIndex];

                    // Only lookups that are, or wrap, a ligature substitution are relevant here.
                    // A tag can legitimately also reference non-ligature lookup types (e.g. a
                    // "liga" record could in principle sit next to unrelated data); those are
                    // simply not collected, exactly like today's UnwrapLigatureSubtable check on
                    // the individual subtables would have discarded them anyway.
                    if (lookup.LookupType != 4 && lookup.LookupType != 7)
                        continue;

                    if (!map.TryGetValue(tag, out var list))
                    {
                        list = new List<IndexedLookup>();
                        map[tag] = list;
                    }

                    list.Add(new IndexedLookup(featureIndex, lookup));
                }
            }

            return map;
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