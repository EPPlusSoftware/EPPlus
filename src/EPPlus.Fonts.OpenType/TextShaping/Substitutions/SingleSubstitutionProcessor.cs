/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/19/2026         EPPlus Software AB           GSUB Single Substitution support
  09/07/2026         EPPlus Software AB           Filter by ScriptList/LangSys, not just tag
  09/07/2026         EPPlus Software AB           Unwrap extension-wrapped (Type 7) lookups
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Substitutions
{
    /// <summary>
    /// Processes GSUB Lookup Type 1 (Single Substitution).
    /// This handles 1:1 glyph replacements like small caps, oldstyle figures, etc.
    /// </summary>
    internal class SingleSubstitutionProcessor
    {
        private readonly struct IndexedSubtable
        {
            public readonly int FeatureIndex;
            public readonly SingleSubstSubTable Subtable;

            public IndexedSubtable(int featureIndex, SingleSubstSubTable subtable)
            {
                FeatureIndex = featureIndex;
                Subtable = subtable;
            }
        }

        private readonly OpenTypeFont _font;
        private readonly GsubTable _gsubTable;

        // Feature tag -> every (FeatureList index, subtable) pair recorded under that tag, across
        // ALL scripts. The FeatureList index lets ApplySubstitutions filter down to the ones the
        // requested script can actually reach - unlike the tag alone, which collapses every
        // script's data for the same tag into one bucket.
        private readonly Dictionary<string, List<IndexedSubtable>> _featureSubtables;
        private readonly Dictionary<string, HashSet<int>> _activeIndexCache = new Dictionary<string, HashSet<int>>();

        public SingleSubstitutionProcessor(OpenTypeFont font)
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
        /// Applies single substitution to the glyph list, restricted to the features reachable
        /// from the given script and language.
        /// This processes all glyphs and replaces them according to the active features.
        /// </summary>
        /// <param name="glyphs">List of shaped glyphs to process</param>
        /// <param name="activeFeatures">List of feature tags to apply (e.g., "smcp", "onum")</param>
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

            // Collect all subtables for the active features
            var subtablesToApply = new List<SingleSubstSubTable>();
            foreach (var feature in activeFeatures)
            {
                if (!_featureSubtables.TryGetValue(feature, out var entries))
                {
                    continue;
                }

                foreach (var entry in entries)
                {
                    // null activeIndices means "no ScriptList to filter by" - keep every entry,
                    // matching the previous behavior rather than discarding features we cannot
                    // resolve.
                    if (activeIndices != null && !activeIndices.Contains(entry.FeatureIndex))
                    {
                        continue;
                    }

                    subtablesToApply.Add(entry.Subtable);
                }
            }

            if (subtablesToApply.Count == 0)
                return glyphs;

            // Process each glyph
            for (int i = 0; i < glyphs.Count; i++)
            {
                ushort originalGlyphId = glyphs[i].GlyphId;

                // Try to find a substitution for this glyph in the active subtables
                if (TryGetSubstitution(originalGlyphId, subtablesToApply, out ushort newGlyphId))
                {
                    // Update both GlyphId and BaseAdvance for the new glyph
                    var glyph = glyphs[i];
                    glyph.GlyphId = newGlyphId;

                    // Get the new glyph's advance width from hmtx
                    var newAdvance = (short)_font.HmtxTable.GetAdvanceWidth(newGlyphId);
                    glyph.BaseAdvance = newAdvance;  // Update base advance
                    glyph.XAdvance = newAdvance;     // Reset to base (kerning will be reapplied)

                    glyphs[i] = glyph;
                }
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
        /// Tries to find a substitution for a given glyph ID in the specified subtables.
        /// </summary>
        private bool TryGetSubstitution(ushort glyphId, List<SingleSubstSubTable> subtables, out ushort substitutedGlyphId)
        {
            substitutedGlyphId = glyphId; // Default to no change

            foreach (var subtable in subtables)
            {
                // Check if this glyph is covered by this subtable
                int coverageIndex = subtable.Coverage?.GetGlyphIndex(glyphId) ?? -1;

                if (coverageIndex >= 0)
                {
                    // Glyph is covered, get the substitution
                    ushort result = subtable.GetSubstitution(glyphId);

                    // Even if result is 0, it's a valid substitution (could be .notdef)
                    // Only skip if it's the same as input (no actual change)
                    if (result != glyphId)
                    {
                        substitutedGlyphId = result;
                        return true;
                    }
                }
            }

            return false; // No substitution found in any subtable
        }

        /// <summary>
        /// Returns the SingleSubstSubTable a GSUB subtable represents, unwrapping Extension
        /// Substitution (Type 7) if needed. Returns null for anything that is not, directly or
        /// via unwrapping, a SingleSubstSubTable.
        /// </summary>
        private static SingleSubstSubTable UnwrapSingleSubstSubtable(Tables.FontTableElement subtableObj)
        {
            if (subtableObj is SingleSubstSubTable direct)
            {
                return direct;
            }

            if (subtableObj is ExtensionSubstSubTable extension
                && extension.ExtensionLookupType == 1
                && extension.ExtendedSubTable is SingleSubstSubTable wrapped)
            {
                return wrapped;
            }

            return null;
        }

        /// <summary>
        /// Builds a map of feature tags to their Single Substitution subtables, keeping each
        /// subtable's original FeatureList index so ApplySubstitutions can filter by script later.
        /// Entries are APPENDED rather than overwritten per tag: two FeatureRecords can
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

                        // Type 1 (Single Substitution) directly, or Type 7 (Extension) wrapping
                        // one. Unlike GPOS, whose loader flattens extension-wrapped lookups so
                        // Type 9 never actually appears on a loaded lookup, GSUB keeps the
                        // wrapper - an extension-wrapped single-substitution lookup has
                        // LookupType 7 with ExtensionSubstSubTable entries in SubTables, each
                        // pointing at the real subtable via ExtendedSubTable.
                        if (lookup.LookupType == 1 || lookup.LookupType == 7)
                        {
                            foreach (var subtableObj in lookup.SubTables)
                            {
                                var singleSubst = UnwrapSingleSubstSubtable(subtableObj);
                                if (singleSubst == null) continue;

                                if (!_featureSubtables.TryGetValue(featureTag, out var list))
                                {
                                    list = new List<IndexedSubtable>();
                                    _featureSubtables[featureTag] = list;
                                }

                                list.Add(new IndexedSubtable(featureIndex, singleSubst));
                            }
                        }
                    }
                }
            }
        }
    }
}