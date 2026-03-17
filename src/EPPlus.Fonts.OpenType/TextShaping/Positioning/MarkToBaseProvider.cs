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
 *************************************************************************************************/
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
        private readonly List<MarkToBaseSubTableFormat1> _subtables;

        public MarkToBaseProvider(OpenTypeFont font)
        {
            if (font.GposTable != null)
            {
                _subtables = FindAllMarkToBaseSubtables(font.GposTable);
            }
            else
            {
                _subtables = new List<MarkToBaseSubTableFormat1>();
            }
        }

        /// <summary>
        /// Applies mark-to-base positioning to a glyph sequence.
        /// Marks are positioned relative to the preceding base glyph.
        /// </summary>
        /// <param name="glyphs">List of shaped glyphs to process</param>
        public void ApplyMarkPositioning(List<ShapedGlyph> glyphs)
        {
            if (_subtables.Count == 0 || glyphs.Count < 2)
                return;

            // Process glyphs left-to-right
            for (int i = 1; i < glyphs.Count; i++)
            {
                var baseGlyph = glyphs[i - 1];
                var markGlyph = glyphs[i];

                bool positioned = false;

                // Try each subtable until we find positioning
                foreach (var subtable in _subtables)
                {
                    if (TryPositionMark(subtable, baseGlyph, markGlyph))
                    {
                        positioned = true;
                        break;
                    }
                }
            }
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

            // Calculate mark position relative to base
            // Mark is positioned so its anchor aligns with base anchor
            var xOffset = baseAnchor.XCoordinate - markAnchor.XCoordinate;
            var yOffset = baseAnchor.YCoordinate - markAnchor.YCoordinate;

            // Apply positioning to mark glyph
            markGlyph.XOffset = (short)xOffset;
            markGlyph.YOffset = (short)yOffset;

            // Mark should not advance (it's positioned over base)
            markGlyph.XAdvance = 0;
            markGlyph.YAdvance = 0;

            return true;
        }

        /// <summary>
        /// Finds all Mark-to-Base subtables in the "mark" feature.
        /// </summary>
        private List<MarkToBaseSubTableFormat1> FindAllMarkToBaseSubtables(GposTable gpos)
        {
            var subtables = new List<MarkToBaseSubTableFormat1>();
            if (gpos == null)
                return subtables;

            foreach (var featureRecord in gpos.FeatureList.FeatureRecords)
            {
                if (featureRecord.FeatureTag.Value == "mark")
                {
                    var feature = featureRecord.FeatureTable;

                    foreach (var lookupIndex in feature.LookupListIndices)
                    {
                        if (lookupIndex >= gpos.LookupList.Lookups.Count)
                            continue;

                        var lookup = gpos.LookupList.Lookups[lookupIndex];

                        // ✅ Kolla bara innehållet, ignorera LookupType
                        foreach (var subtable in lookup.SubTables)
                        {
                            if (subtable is MarkToBaseSubTableFormat1 markToBase)
                            {
                                subtables.Add(markToBase);
                            }
                        }
                    }
                }
            }

            return subtables;
        }
    }
}