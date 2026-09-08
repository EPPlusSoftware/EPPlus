/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           Mark-to-Base positioning
  09/07/2026         EPPlus Software AB           Pen-relative, additive mark offsets
  09/07/2026         EPPlus Software AB           Filter by ScriptList/LangSys, not just tag
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Positioning
{
    /// <summary>
    /// Provides mark-to-base attachment positioning (GPOS Type 4).
    /// Positions combining marks (accents, diacritics) relative to base glyphs.
    /// Critical for decomposed Unicode text (e.g., e + ´ → é).
    /// </summary>
    internal class MarkToBaseProvider
    {
        private readonly struct IndexedSubtable
        {
            public readonly int FeatureIndex;
            public readonly MarkToBaseSubTableFormat1 Subtable;

            public IndexedSubtable(int featureIndex, MarkToBaseSubTableFormat1 subtable)
            {
                FeatureIndex = featureIndex;
                Subtable = subtable;
            }
        }

        private readonly GposTable _gpos;
        private readonly List<IndexedSubtable> _subtables;
        private readonly Dictionary<string, HashSet<int>> _activeIndexCache = new Dictionary<string, HashSet<int>>();

        public MarkToBaseProvider(OpenTypeFont font)
        {
            _gpos = font.GposTable;

            if (_gpos != null)
            {
                _subtables = FindAllMarkToBaseSubtables(_gpos);
            }
            else
            {
                _subtables = new List<IndexedSubtable>();
            }
        }

        /// <summary>
        /// Applies mark-to-base positioning to a glyph sequence.
        /// Marks are positioned relative to the preceding base glyph.
        /// </summary>
        /// <param name="glyphs">List of shaped glyphs to process</param>
        /// <param name="script">
        /// OpenType script tag (e.g. "latn"). Pass null to fall back to unfiltered lookup, which
        /// reproduces the previous behavior for callers that have no script to give.
        /// </param>
        /// <param name="language">OpenType language-system tag, or null for the script's default.</param>
        public void ApplyMarkPositioning(List<ShapedGlyph> glyphs, string script, string language)
        {
            if (_subtables.Count == 0 || glyphs.Count < 2)
                return;

            HashSet<int> activeIndices = GetActiveIndices(script, language);

            // Process glyphs left-to-right
            for (int i = 1; i < glyphs.Count; i++)
            {
                var baseGlyph = glyphs[i - 1];
                var markGlyph = glyphs[i];

                // Try each subtable reachable from the active script until we find positioning
                foreach (var entry in _subtables)
                {
                    // null activeIndices means "no ScriptList to filter by" - keep every entry,
                    // matching the previous behavior rather than discarding marks we cannot resolve.
                    if (activeIndices != null && !activeIndices.Contains(entry.FeatureIndex))
                    {
                        continue;
                    }

                    if (TryPositionMark(entry.Subtable, baseGlyph, markGlyph))
                    {
                        break;
                    }
                }
            }
        }

        private HashSet<int> GetActiveIndices(string script, string language)
        {
            string cacheKey = (script ?? string.Empty) + "|" + (language ?? string.Empty);

            if (!_activeIndexCache.TryGetValue(cacheKey, out var indices))
            {
                indices = ScriptFeatureResolver.GetActiveFeatureIndices(_gpos?.ScriptList, script, language);
                _activeIndexCache[cacheKey] = indices;
            }

            return indices;
        }

        /// <summary>
        /// Attempts to position a mark glyph relative to a base glyph.
        /// </summary>
        private bool TryPositionMark(
            MarkToBaseSubTableFormat1 subtable,
            ShapedGlyph baseGlyph,
            ShapedGlyph markGlyph)
        {
            // Check if base glyph is in base coverage
            int baseIndex = subtable.BaseCoverage?.GetGlyphIndex(baseGlyph.GlyphId) ?? -1;
            if (baseIndex < 0 || baseIndex >= subtable.BaseArray.BaseCount)
                return false;

            // Check if mark glyph is in mark coverage
            int markIndex = subtable.MarkCoverage?.GetGlyphIndex(markGlyph.GlyphId) ?? -1;
            if (markIndex < 0 || markIndex >= subtable.MarkArray.MarkCount)
                return false;

            // Get mark record (contains class and anchor)
            var markRecord = subtable.MarkArray.Records[markIndex];
            ushort markClass = markRecord.MarkClass;

            // Validate mark class
            if (markClass >= subtable.MarkClassCount)
                return false;

            // Get base record (contains anchors for each mark class)
            var baseRecord = subtable.BaseArray.Records[baseIndex];
            if (baseRecord.BaseAnchors == null || markClass >= baseRecord.BaseAnchors.Length)
                return false;

            var baseAnchor = baseRecord.BaseAnchors[markClass];
            var markAnchor = markRecord.MarkAnchor;

            if (baseAnchor == null || markAnchor == null)
                return false;

            // Calculate mark position relative to the PEN, not to the base glyph's origin
            // (HarfBuzz convention). The pen has already moved by the base glyph's advance when
            // the mark is drawn, so that advance has to be taken back out of the offset.
            // baseGlyph.XAdvance is used rather than the raw hmtx advance because kerning has
            // already been applied at this point - ApplyPositioning orders SinglePos, then
            // kerning, then mark positioning.
            var xOffset = baseAnchor.XCoordinate - markAnchor.XCoordinate - baseGlyph.XAdvance;
            var yOffset = baseAnchor.YCoordinate - markAnchor.YCoordinate;

            // Accumulate rather than assign, so an XPlacement/YPlacement already written by
            // SinglePos (TextShaper, ApplyValueRecord) is not discarded.
            markGlyph.XOffset += (short)xOffset;
            markGlyph.YOffset += (short)yOffset;

            // Mark should not advance (it's positioned over base)
            markGlyph.XAdvance = 0;
            markGlyph.YAdvance = 0;

            return true;
        }

        /// <summary>
        /// Finds all Mark-to-Base subtables in "mark"-tagged FeatureRecords, together with each
        /// one's original FeatureList index, across all scripts. Two FeatureRecords can
        /// legitimately share the "mark" tag (one per script); both are kept so
        /// ApplyMarkPositioning can filter by script at call time.
        /// </summary>
        private List<IndexedSubtable> FindAllMarkToBaseSubtables(GposTable gpos)
        {
            var subtables = new List<IndexedSubtable>();

            if (gpos?.FeatureList == null)
                return subtables;

            var featureRecords = gpos.FeatureList.FeatureRecords;

            for (int featureIndex = 0; featureIndex < featureRecords.Count; featureIndex++)
            {
                var featureRecord = featureRecords[featureIndex];

                if (featureRecord.FeatureTag.Value != "mark")
                    continue;

                var feature = featureRecord.FeatureTable;

                foreach (var lookupIndex in feature.LookupListIndices)
                {
                    if (lookupIndex >= gpos.LookupList.Lookups.Count)
                        continue;

                    var lookup = gpos.LookupList.Lookups[lookupIndex];

                    // Only the subtable content is checked, not LookupType - extension wrapped
                    // (type 9) mark lookups are unwrapped by GposTableLoader but keep LookupType 9.
                    foreach (var subtable in lookup.SubTables)
                    {
                        if (subtable is MarkToBaseSubTableFormat1 markToBase)
                        {
                            subtables.Add(new IndexedSubtable(featureIndex, markToBase));
                        }
                    }
                }
            }

            return subtables;
        }
    }
}