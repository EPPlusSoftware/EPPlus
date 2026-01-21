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
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.TextShaping.Substitutions
{
    /// <summary>
    /// Processes GSUB Lookup Type 1 (Single Substitution).
    /// This handles 1:1 glyph replacements like small caps, oldstyle figures, etc.
    /// </summary>
    internal class SingleSubstitutionProcessor
    {
        private readonly GsubTable _gsubTable;
        private readonly Dictionary<string, List<SingleSubstSubTable>> _featureSubtables;

        public SingleSubstitutionProcessor(OpenTypeFont font)
        {
            _gsubTable = font?.GsubTable;
            _featureSubtables = new Dictionary<string, List<SingleSubstSubTable>>();

            if (_gsubTable != null)
            {
                BuildFeatureSubtableMap();
            }
        }

        /// <summary>
        /// Applies single substitution to the glyph list.
        /// This processes all glyphs and replaces them according to the active features.
        /// </summary>
        /// <param name="glyphs">List of shaped glyphs to process</param>
        /// <param name="activeFeatures">List of feature tags to apply (e.g., "smcp", "onum")</param>
        /// <returns>Modified glyph list with substitutions applied</returns>
        public List<ShapedGlyph> ApplySubstitutions(List<ShapedGlyph> glyphs, List<string> activeFeatures)
        {
            if (glyphs == null || glyphs.Count == 0)
                return glyphs;

            if (activeFeatures == null || activeFeatures.Count == 0)
                return glyphs;

            // Collect all subtables for the active features
            var subtablesToApply = new List<SingleSubstSubTable>();
            foreach (var feature in activeFeatures)
            {
                if (_featureSubtables.TryGetValue(feature, out var subtables))
                {
                    subtablesToApply.AddRange(subtables);
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
                    // Replace the glyph ID while keeping all other properties
                    var glyph = glyphs[i];
                    glyph.GlyphId = newGlyphId;
                    glyphs[i] = glyph;
                }
            }

            return glyphs;
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
        /// Builds a map of feature tags to their Single Substitution subtables.
        /// This allows us to only apply substitutions for active features.
        /// </summary>
        private void BuildFeatureSubtableMap()
        {
            if (_gsubTable?.FeatureList == null || _gsubTable.LookupList == null)
                return;

            foreach (var featureRecord in _gsubTable.FeatureList.FeatureRecords)
            {
                string featureTag = featureRecord.FeatureTag.Value;
                var feature = featureRecord.FeatureTable;

                if (feature?.LookupListIndices == null)
                    continue;

                // Get or create the list for this feature
                if (!_featureSubtables.ContainsKey(featureTag))
                {
                    _featureSubtables[featureTag] = new List<SingleSubstSubTable>();
                }

                var subtables = _featureSubtables[featureTag];

                // Get all Single Substitution subtables for this feature
                foreach (var lookupIndex in feature.LookupListIndices)
                {
                    if (lookupIndex >= _gsubTable.LookupList.Lookups.Count)
                        continue;

                    var lookup = _gsubTable.LookupList.Lookups[lookupIndex];

                    // We want Single Substitution (Type 1)
                    if (lookup.LookupType == 1 && lookup.SubTables != null)
                    {
                        foreach (var subtable in lookup.SubTables)
                        {
                            if (subtable is SingleSubstSubTable singleSubst)
                            {
                                subtables.Add(singleSubst);
                            }
                        }
                    }
                }
            }
        }
    }
}