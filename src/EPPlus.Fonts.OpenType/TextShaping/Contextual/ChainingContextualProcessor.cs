/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.TextShaping.Ligatures;
using EPPlus.Fonts.OpenType.TextShaping.Substitutions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping.Contextual
{
    internal class ChainingContextualProcessor
    {
        private readonly OpenTypeFont _font;
        private readonly SingleSubstitutionProcessor _singleSubstProcessor;
        private readonly LigatureProcessor _ligatureProcessor;

        public ChainingContextualProcessor(
            OpenTypeFont font,
            SingleSubstitutionProcessor singleSubstProcessor,
            LigatureProcessor ligatureProcessor)
        {
            _font = font;
            _singleSubstProcessor = singleSubstProcessor;
            _ligatureProcessor = ligatureProcessor;
        }

        /// <summary>
        /// Applies chaining contextual substitutions for a specific feature.
        /// </summary>
        internal List<ShapedGlyph> ApplyContextualSubstitutions(
            List<ShapedGlyph> glyphs,
            string featureTag)
        {
            var gsub = _font.GsubTable;
            if (gsub == null)
                return glyphs;

            // Find all Type 6 lookups for this feature
            var contextualLookups = FindContextualLookupsForFeature(gsub, featureTag);
            if (contextualLookups.Count == 0)
                return glyphs;

            // Apply each lookup in order
            foreach (var lookup in contextualLookups)
            {
                glyphs = ApplyContextualLookup(glyphs, lookup);
            }

            return glyphs;
        }

        /// <summary>
        /// Finds all Type 6 lookups associated with a feature tag.
        /// </summary>
        private List<LookupTable> FindContextualLookupsForFeature(GsubTable gsub, string featureTag)
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
                            var lookup = gsub.LookupList.Lookups[lookupIndex];

                            // Only Type 6 (Chaining Contextual) or Type 7 (Extension wrapping Type 6)
                            if (lookup.LookupType == 6)
                            {
                                lookups.Add(lookup);
                            }
                            else if (lookup.LookupType == 7)
                            {
                                // Check if extension wraps a Type 6
                                foreach (var subtable in lookup.SubTables)
                                {
                                    if (subtable is ExtensionSubstSubTable ext &&
                                        ext.ExtensionLookupType == 6)
                                    {
                                        lookups.Add(lookup);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return lookups;
        }

        /// <summary>
        /// Applies a single chaining contextual lookup to the glyph sequence.
        /// </summary>
        private List<ShapedGlyph> ApplyContextualLookup(List<ShapedGlyph> glyphs, LookupTable lookup)
        {
            var result = new List<ShapedGlyph>(glyphs);
            int i = 0;

            while (i < result.Count)
            {
                bool substituted = false;

                // Try each subtable
                foreach (var subtable in lookup.SubTables)
                {
                    ChainingContextualSubstFormat3 contextual = null;

                    if (subtable is ChainingContextualSubstFormat3 format3)
                    {
                        contextual = format3;
                    }
                    else if (subtable is ExtensionSubstSubTable ext &&
                             ext.ExtendedSubTable is ChainingContextualSubstFormat3 extFormat3)
                    {
                        contextual = extFormat3;
                    }

                    if (contextual != null)
                    {
                        // Try to match and apply contextual rule at position i
                        if (TryApplyContextualRule(result, i, contextual, out var newGlyphs, out int glyphsConsumed))
                        {
                            // Replace glyphs at position i with the result
                            result.RemoveRange(i, glyphsConsumed);
                            result.InsertRange(i, newGlyphs);

                            // Move past the substituted sequence
                            i += newGlyphs.Count;
                            substituted = true;
                            break;
                        }
                    }
                }

                if (!substituted)
                {
                    i++;
                }
            }

            return result;
        }

        /// <summary>
        /// Attempts to match and apply a contextual rule starting at the given position.
        /// </summary>
        private bool TryApplyContextualRule(
            List<ShapedGlyph> glyphs,
            int position,
            ChainingContextualSubstFormat3 rule,
            out List<ShapedGlyph> resultGlyphs,
            out int glyphsConsumed)
        {
            resultGlyphs = null;
            glyphsConsumed = 0;

            // 1. Check if we have enough glyphs for the complete context
            int backtrackCount = rule.BacktrackCoverages?.Count ?? 0;
            int inputCount = rule.InputCoverages?.Count ?? 0;
            int lookaheadCount = rule.LookaheadCoverages?.Count ?? 0;

            if (inputCount == 0)
                return false;

            // Check bounds
            if (position < backtrackCount)
                return false;

            if (position + inputCount + lookaheadCount > glyphs.Count)
                return false;

            // 2. Match backtrack context (in reverse order!)
            if (backtrackCount > 0)
            {
                for (int i = 0; i < backtrackCount; i++)
                {
                    int glyphPos = position - 1 - i;
                    var coverage = rule.BacktrackCoverages[i];

                    if (coverage.GetGlyphIndex(glyphs[glyphPos].GlyphId) < 0)
                        return false; // Backtrack mismatch
                }
            }

            // 3. Match input sequence
            for (int i = 0; i < inputCount; i++)
            {
                int glyphPos = position + i;
                var coverage = rule.InputCoverages[i];

                if (coverage.GetGlyphIndex(glyphs[glyphPos].GlyphId) < 0)
                    return false; // Input mismatch
            }

            // 4. Match lookahead context
            if (lookaheadCount > 0)
            {
                for (int i = 0; i < lookaheadCount; i++)
                {
                    int glyphPos = position + inputCount + i;
                    var coverage = rule.LookaheadCoverages[i];

                    if (coverage.GetGlyphIndex(glyphs[glyphPos].GlyphId) < 0)
                        return false; // Lookahead mismatch
                }
            }

            // 5. Context matches! Apply the substitution lookups
            var inputGlyphs = glyphs.GetRange(position, inputCount);

            foreach (var substRecord in rule.SubstLookupRecords)
            {
                if (substRecord.SequenceIndex >= inputCount)
                    continue; // Invalid record

                // Get the lookup to apply
                var gsub = _font.GsubTable;
                if (substRecord.LookupListIndex >= gsub.LookupList.Lookups.Count)
                    continue;

                var targetLookup = gsub.LookupList.Lookups[substRecord.LookupListIndex];

                // Apply the lookup to the input sequence
                inputGlyphs = ApplyReferencedLookup(inputGlyphs, targetLookup, substRecord.SequenceIndex);
            }

            resultGlyphs = inputGlyphs;
            glyphsConsumed = inputCount;
            return true;
        }

        /// <summary>
        /// Applies a referenced lookup (Type 1, Type 4, etc.) to a glyph sequence.
        /// </summary>
        private List<ShapedGlyph> ApplyReferencedLookup(
            List<ShapedGlyph> glyphs,
            LookupTable lookup,
            int startPosition)
        {
            // Get the actual lookup type (unwrap Extension if needed)
            ushort lookupType = lookup.LookupType;
            List<FontTableElement> subtables = lookup.SubTables;

            if (lookupType == 7 && subtables.Count > 0 && subtables[0] is ExtensionSubstSubTable ext)
            {
                lookupType = ext.ExtensionLookupType;
                subtables = new List<FontTableElement> { ext.ExtendedSubTable };
            }

            switch (lookupType)
            {
                case 1: // Single Substitution
                    return ApplySingleSubstitutionAtPosition(glyphs, subtables, startPosition);

                case 4: // Ligature Substitution
                    return ApplyLigatureSubstitutionAtPosition(glyphs, subtables, startPosition);

                default:
                    // Unsupported lookup type in contextual rule
                    return glyphs;
            }
        }

        /// <summary>
        /// Applies Type 1 Single Substitution at a specific position.
        /// </summary>
        private List<ShapedGlyph> ApplySingleSubstitutionAtPosition(
            List<ShapedGlyph> glyphs,
            List<FontTableElement> subtables,
            int position)
        {
            if (position >= glyphs.Count)
                return glyphs;

            foreach (var subtable in subtables)
            {
                if (subtable is SingleSubstSubTable singleSubst)
                {
                    ushort oldGlyphId = glyphs[position].GlyphId;

                    var newGlyphId = singleSubst.GetSubstitution(oldGlyphId);
                    if (newGlyphId > 0)
                    {
                        var result = new List<ShapedGlyph>(glyphs);
                        result[position] = CreateSubstitutedGlyph(glyphs[position], newGlyphId);
                        return result;
                    }
                }
            }

            return glyphs;
        }

        /// <summary>
        /// Applies Type 4 Ligature Substitution starting at a specific position.
        /// </summary>
        private List<ShapedGlyph> ApplyLigatureSubstitutionAtPosition(
            List<ShapedGlyph> glyphs,
            List<FontTableElement> subtables,
            int position)
        {
            if (position >= glyphs.Count)
                return glyphs;

            foreach (var subtable in subtables)
            {
                if (subtable is LigatureSubstSubTable ligSubtable)
                {
                    ushort firstGlyph = glyphs[position].GlyphId;

                    // Check coverage
                    int coverageIndex = ligSubtable.Coverage.GetGlyphIndex(firstGlyph);
                    if (coverageIndex < 0)
                        continue;

                    if (!ligSubtable.LigatureSets.TryGetValue(firstGlyph, out var ligatureSet))
                        continue;

                    if (ligatureSet?.Ligatures == null)
                        continue;

                    // Try each ligature
                    foreach (var ligature in ligatureSet.Ligatures)
                    {
                        int componentCount = 1 + (ligature.Components?.Length ?? 0);

                        if (position + componentCount > glyphs.Count)
                            continue;

                        // Check if components match
                        bool matches = true;
                        if (ligature.Components != null)
                        {
                            for (int i = 0; i < ligature.Components.Length; i++)
                            {
                                if (glyphs[position + 1 + i].GlyphId != ligature.Components[i])
                                {
                                    matches = false;
                                    break;
                                }
                            }
                        }

                        if (matches)
                        {
                            // Create ligature
                            var result = new List<ShapedGlyph>(glyphs);
                            var ligatureGlyph = CreateLigatureGlyph(glyphs, position, (byte)componentCount, ligature.LigatureGlyph);

                            result.RemoveRange(position, componentCount);
                            result.Insert(position, ligatureGlyph);

                            return result;
                        }
                    }
                }
            }

            return glyphs;
        }

        private ShapedGlyph CreateSubstitutedGlyph(ShapedGlyph original, ushort newGlyphId)
        {
            var baseAdvance = (short)_font.HmtxTable.GetAdvanceWidth(newGlyphId);

            return new ShapedGlyph
            {
                GlyphId = newGlyphId,
                BaseAdvance = baseAdvance,      // ← New base advance for substituted glyph
                XAdvance = baseAdvance,         // ← Reset to base (kerning will be reapplied)
                YAdvance = 0,
                XOffset = 0,
                YOffset = 0,
                ClusterIndex = original.ClusterIndex,
                CharCount = original.CharCount
            };
        }

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
                BaseAdvance = baseAdvance,      // ← Base advance for ligature
                XAdvance = baseAdvance,         // ← Will be adjusted by positioning
                YAdvance = 0,
                XOffset = 0,
                YOffset = 0,
                ClusterIndex = clusterIndex,
                CharCount = componentCount
            };
        }
    }
}