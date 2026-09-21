/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution (Type 2) support
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Substitutions
{
    /// <summary>
    /// Processes GSUB Lookup Type 2 (Multiple Substitution).
    /// This handles 1:N glyph expansions - the opposite direction of ligature substitution
    /// (Type 4). Fonts commonly use it for the "ccmp" feature to decompose a precomposed glyph
    /// into a base glyph plus combining marks.
    /// </summary>
    internal class MultipleSubstitutionProcessor
    {
        private readonly struct IndexedSubtable
        {
            public readonly int FeatureIndex;
            public readonly MultipleSubstSubTable Subtable;

            public IndexedSubtable(int featureIndex, MultipleSubstSubTable subtable)
            {
                FeatureIndex = featureIndex;
                Subtable = subtable;
            }
        }

        private readonly OpenTypeFont _font;
        private readonly GsubTable _gsubTable;

        // Feature tag -> every (FeatureList index, subtable) pair recorded under that tag, across
        // ALL scripts - same shape as SingleSubstitutionProcessor's map, for the same reason: the
        // FeatureList index lets ApplySubstitutions filter down to what the requested script can
        // actually reach.
        private readonly Dictionary<string, List<IndexedSubtable>> _featureSubtables;
        private readonly Dictionary<string, HashSet<int>> _activeIndexCache = new Dictionary<string, HashSet<int>>();

        public MultipleSubstitutionProcessor(OpenTypeFont font)
        {
            _font = font;
            _gsubTable = font?.GsubTable;
            _featureSubtables = new Dictionary<string, List<IndexedSubtable>>();

            if (_gsubTable != null)
            {
                BuildFeatureSubtableMap();
            }
        }

        /// <summary>
        /// Applies multiple substitution to the glyph list, restricted to the features reachable
        /// from the given script and language. A matched glyph is REPLACED by its substitute
        /// sequence in place, so the returned list can be longer than the input.
        /// </summary>
        /// <param name="glyphs">List of shaped glyphs to process</param>
        /// <param name="activeFeatures">List of feature tags to apply (e.g., "ccmp")</param>
        /// <param name="script">
        /// OpenType script tag (e.g. "latn"). Pass null to fall back to unfiltered lookup, which
        /// reproduces the previous behavior for callers that have no script to give.
        /// </param>
        /// <param name="language">OpenType language-system tag, or null for the script's default.</param>
        /// <returns>Modified glyph list with substitutions applied</returns>
        public List<ShapedGlyph> ApplySubstitutions(List<ShapedGlyph> glyphs, List<string> activeFeatures, string script, string language)
        {
            if (glyphs == null || glyphs.Count == 0)
                return glyphs;

            if (activeFeatures == null || activeFeatures.Count == 0)
                return glyphs;

            HashSet<int> activeIndices = GetActiveIndices(script, language);

            var subtablesToApply = new List<MultipleSubstSubTable>();
            foreach (var feature in activeFeatures)
            {
                if (!_featureSubtables.TryGetValue(feature, out var entries))
                {
                    continue;
                }

                foreach (var entry in entries)
                {
                    if (activeIndices != null && !activeIndices.Contains(entry.FeatureIndex))
                    {
                        continue;
                    }

                    subtablesToApply.Add(entry.Subtable);
                }
            }

            if (subtablesToApply.Count == 0)
                return glyphs;

            for (int i = 0; i < glyphs.Count; i++)
            {
                ushort[] sequence = TryGetSequence(glyphs[i].GlyphId, subtablesToApply);
                if (sequence == null || sequence.Length == 0)
                    continue;

                ShapedGlyph original = glyphs[i];
                glyphs.RemoveAt(i);

                var newGlyphs = new List<ShapedGlyph>(sequence.Length);
                for (int s = 0; s < sequence.Length; s++)
                {
                    ushort newGlyphId = sequence[s];
                    var advance = (short)_font.HmtxTable.GetAdvanceWidth(newGlyphId);

                    newGlyphs.Add(new ShapedGlyph
                    {
                        GlyphId = newGlyphId,
                        BaseAdvance = advance,
                        XAdvance = advance,
                        YAdvance = 0,
                        XOffset = 0,
                        YOffset = 0,
                        ClusterIndex = original.ClusterIndex,
                        // The whole original CharCount belongs to the first output glyph, so the
                        // total across the expanded glyphs still adds up to the source text
                        // length - the same convention LigatureProcessor uses in reverse (there,
                        // several input CharCounts are summed onto the one output glyph).
                        CharCount = (byte)(s == 0 ? original.CharCount : 0),
                        FontId = original.FontId
                    });
                }

                glyphs.InsertRange(i, newGlyphs);
                i += newGlyphs.Count - 1; // Skip past the glyphs just inserted.
            }

            return glyphs;
        }

        private HashSet<int> GetActiveIndices(string script, string language)
        {
            string cacheKey = (script ?? string.Empty) + "|" + (language ?? string.Empty);

            if (!_activeIndexCache.TryGetValue(cacheKey, out var indices))
            {
                indices = ScriptFeatureResolver.GetActiveFeatureIndices(_gsubTable?.ScriptList, script, language);
                _activeIndexCache[cacheKey] = indices;
            }

            return indices;
        }

        /// <summary>
        /// Tries to find a substitute sequence for a given glyph ID in the specified subtables.
        /// </summary>
        private static ushort[] TryGetSequence(ushort glyphId, List<MultipleSubstSubTable> subtables)
        {
            foreach (var subtable in subtables)
            {
                int coverageIndex = subtable.Coverage?.GetGlyphIndex(glyphId) ?? -1;
                if (coverageIndex < 0)
                    continue;

                ushort[] sequence = subtable.GetSubstitution(glyphId);
                if (sequence != null)
                {
                    return sequence;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns the MultipleSubstSubTable a GSUB subtable represents, unwrapping Extension
        /// Substitution (Type 7) if needed. Returns null for anything that is not, directly or
        /// via unwrapping, a MultipleSubstSubTable.
        /// </summary>
        private static MultipleSubstSubTable UnwrapMultipleSubstSubtable(Tables.FontTableElement subtableObj)
        {
            if (subtableObj is MultipleSubstSubTable direct)
            {
                return direct;
            }

            if (subtableObj is ExtensionSubstSubTable extension
                && extension.ExtensionLookupType == 2
                && extension.ExtendedSubTable is MultipleSubstSubTable wrapped)
            {
                return wrapped;
            }

            return null;
        }

        /// <summary>
        /// Builds a map of feature tags to their Multiple Substitution subtables, keeping each
        /// subtable's original FeatureList index so ApplySubstitutions can filter by script
        /// later. Entries are APPENDED rather than overwritten per tag: two FeatureRecords can
        /// legitimately share a tag (one per script), and both must survive so the script filter
        /// has something to choose between at lookup time.
        /// </summary>
        private void BuildFeatureSubtableMap()
        {
            if (_gsubTable?.FeatureList?.FeatureRecords == null)
                return;

            var featureRecords = _gsubTable.FeatureList.FeatureRecords;

            for (int featureIndex = 0; featureIndex < featureRecords.Count; featureIndex++)
            {
                var featureRecord = featureRecords[featureIndex];
                string featureTag = featureRecord.FeatureTag.Value;

                var feature = featureRecord.FeatureTable;

                foreach (var lookupIndex in feature.LookupListIndices)
                {
                    if (lookupIndex < _gsubTable.LookupList.Lookups.Count)
                    {
                        var lookup = _gsubTable.LookupList.Lookups[lookupIndex];

                        // Type 2 (Multiple Substitution) directly, or Type 7 (Extension) wrapping
                        // one - same reasoning as SingleSubstitutionProcessor: GSUB keeps the
                        // wrapper rather than flattening it at load time.
                        if (lookup.LookupType == 2 || lookup.LookupType == 7)
                        {
                            foreach (var subtableObj in lookup.SubTables)
                            {
                                var multiSubst = UnwrapMultipleSubstSubtable(subtableObj);
                                if (multiSubst == null) continue;

                                if (!_featureSubtables.TryGetValue(featureTag, out var list))
                                {
                                    list = new List<IndexedSubtable>();
                                    _featureSubtables[featureTag] = list;
                                }

                                list.Add(new IndexedSubtable(featureIndex, multiSubst));
                            }
                        }
                    }
                }
            }
        }
    }
}