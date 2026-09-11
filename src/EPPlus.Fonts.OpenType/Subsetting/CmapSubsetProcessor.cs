/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class CmapSubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            // --- PHASE 1: DISCOVERY ---
            // In this phase, we only find the original Glyph IDs for the requested characters.
            // These will be added to context.IncludedGlyphs so that GlyfAndLocaProcessor
            // knows which glyph data to copy.
            foreach (uint codePoint in context.UsedCodePoints)
            {
                ushort oldGid;
                if (context.OriginalFont.CmapTable.TryGetGlyphId(codePoint, out oldGid))
                {
                    if (!context.IncludedGlyphs.Contains(oldGid))
                    {
                        context.IncludedGlyphs.Add(oldGid);
                    }
                }
            }

            // --- Unicode Variation Sequences (cmap format 14) ---
            // context.UsedCodePoints is a flat set of code points with no notion of "this selector
            // followed this base char in the text" - CodePointUtil.ExtractCodePoints just decodes
            // UTF-16 to scalar values, so a base char and a variation selector that appeared
            // together in the text are indistinguishable here from two that never did. So for every
            // variation selector actually present in the used code points, every OTHER used code
            // point is checked against the original font's format-14 table, and any pair that IS
            // registered there is kept. A pair that's registered in the font but never actually
            // adjacent in the real text is a harmless false positive - a few extra bytes/glyphs in
            // the subset - because TextShaper only ever looks up a pair it finds truly adjacent, so
            // an over-included pair is simply never queried at render time.
            var subtable14 = FindFormat14Subtable(context.OriginalFont);
            if (subtable14 != null)
            {
                foreach (var selector in subtable14.VariationSelectors)
                {
                    if (!context.UsedCodePoints.Contains(selector.VarSelector))
                        continue;

                    foreach (uint baseCodePoint in context.UsedCodePoints)
                    {
                        if (baseCodePoint == selector.VarSelector)
                            continue;

                        ushort variantGid;
                        if (context.OriginalFont.CmapTable.TryGetGlyphId(baseCodePoint, selector.VarSelector, out variantGid))
                        {
                            if (!context.IncludedGlyphs.Contains(variantGid))
                            {
                                context.IncludedGlyphs.Add(variantGid);
                            }
                        }
                    }
                }
            }

            // Ensure GID 0 (.notdef) is always included
            if (!context.IncludedGlyphs.Contains(0))
            {
                context.IncludedGlyphs.Add(0);
            }
        }

        public void Rewrite(FontSubsettingContext context)
        {
            // Build mapping: Unicode code point → NEW glyph ID in subset
            Dictionary<uint, ushort> cmapMapping = new Dictionary<uint, ushort>();

            foreach (uint codePoint in context.UsedCodePoints)
            {
                ushort oldGid;
                if (context.OriginalFont.CmapTable.TryGetGlyphId(codePoint, out oldGid))
                {
                    ushort newGid;
                    if (context.OldToNewGlyphId.TryGetValue(oldGid, out newGid))
                    {
                        cmapMapping[codePoint] = newGid;
                    }
                    else
                    {
                        cmapMapping[codePoint] = 0; // .notdef fallback
                    }
                }
            }

            // Always map code point 0 to .notdef (required by spec)
            cmapMapping[0] = 0;

            // Check if we need Format 12 (for code points > 0xFFFF like emoji)
            bool needsFormat12 = cmapMapping.Keys.Any(cp => cp > 0xFFFF);

            // Build new cmap table
            CmapTable newCmap = new CmapTable();
            newCmap.Version = 0;

            if (needsFormat12)
            {
                // Create Format 12 subtable for full Unicode support
                var format12 = CreateFormat12Subtable(cmapMapping);

                // Also create Format 4 for BMP characters (backwards compatibility)
                var bmpMapping = cmapMapping.Where(kvp => kvp.Key <= 0xFFFF)
                                             .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                var format4 = CmapFormat4.CreateFromMappings(bmpMapping);

                // Add Format 12 record (3,10) – Windows Unicode UCS-4 (full range)
                EncodingRecord format12Record = new EncodingRecord(Platforms.Windows, 10, 0);
                format12Record.Subtable = format12;
                newCmap.EncodingRecords.Add(format12Record);
                newCmap.SubTables.Add(format12);

                // Add Format 4 record (3,1) – Windows Unicode BMP (backwards compatibility)
                EncodingRecord format4Record = new EncodingRecord(Platforms.Windows, 1, 0);
                format4Record.Subtable = format4;
                newCmap.EncodingRecords.Add(format4Record);
                newCmap.SubTables.Add(format4);

                newCmap.NumTables = 2;
            }
            else
            {
                // Only BMP characters - Format 4 is sufficient
                CmapSubtable4 format4 = CmapFormat4.CreateFromMappings(cmapMapping);

                // (3,1) – Windows Unicode BMP
                EncodingRecord winRecord = new EncodingRecord(Platforms.Windows, 1, 0);
                winRecord.Subtable = format4;

                // (0,3) – Unicode BMP
                EncodingRecord unicodeRecord = new EncodingRecord(Platforms.Unicode, 3, 0);
                unicodeRecord.Subtable = format4;

                newCmap.EncodingRecords.Add(winRecord);
                newCmap.EncodingRecords.Add(unicodeRecord);
                newCmap.SubTables.Add(format4);
                newCmap.NumTables = 2;
            }

            // --- Unicode Variation Sequences (cmap format 14) ---
            // Preserve the (base, selector) pairs that Discover found registered in the original
            // font and that are actually used in this subset, remapped to the subset's new glyph
            // IDs. Without this, TextShaper's format-14 lookahead (which runs against whichever
            // font it's actually shaping - full or subset) would silently fall back to the base
            // character's default glyph when shaping against the embedded subset, even though the
            // full font correctly picked a variant.
            var originalSubtable14 = FindFormat14Subtable(context.OriginalFont);
            if (originalSubtable14 != null)
            {
                var newSubtable14 = BuildSubsetFormat14Subtable(originalSubtable14, context);
                if (newSubtable14 != null)
                {
                    // (0,5) - Unicode Variation Sequences: the platform/encoding combination the
                    // OpenType spec registers for format 14.
                    EncodingRecord uvsRecord = new EncodingRecord(Platforms.Unicode, 5, 0);
                    uvsRecord.Subtable = newSubtable14;
                    newCmap.EncodingRecords.Add(uvsRecord);
                    newCmap.SubTables.Add(newSubtable14);
                    newCmap.NumTables++;
                }
            }

            context.SubsetFont.AddOrReplaceTable(newCmap);
        }

        private static CmapSubtable14 FindFormat14Subtable(OpenTypeFont font)
        {
            foreach (var subtable in font.CmapTable.SubTables)
            {
                if (subtable.Format == 14)
                    return subtable as CmapSubtable14;
            }
            return null;
        }

        /// <summary>
        /// Rebuilds a format-14 subtable containing only the variation selectors, base characters
        /// and glyph IDs that are both registered in <paramref name="original"/> AND actually
        /// present in this subset's used code points / retained glyph mapping. Returns null if
        /// nothing survives the filter (e.g. the text used no variation sequences at all).
        /// </summary>
        private CmapSubtable14 BuildSubsetFormat14Subtable(CmapSubtable14 original, FontSubsettingContext context)
        {
            var newSubtable14 = new CmapSubtable14();

            foreach (var selector in original.VariationSelectors)
            {
                if (!context.UsedCodePoints.Contains(selector.VarSelector))
                    continue;

                NonDefaultUvsTable newNonDefault = null;
                if (selector.NonDefaultUvsTable != null)
                {
                    foreach (var mapping in selector.NonDefaultUvsTable.Mappings)
                    {
                        ushort newGid;
                        if (context.UsedCodePoints.Contains(mapping.UnicodeValue) &&
                            context.OldToNewGlyphId.TryGetValue(mapping.GlyphId, out newGid))
                        {
                            if (newNonDefault == null)
                                newNonDefault = new NonDefaultUvsTable { Mappings = new List<UvsMapping>() };

                            newNonDefault.Mappings.Add(new UvsMapping { UnicodeValue = mapping.UnicodeValue, GlyphId = newGid });
                        }
                    }
                }

                DefaultUvsTable newDefault = null;
                if (selector.DefaultUvsTable != null)
                {
                    foreach (var range in selector.DefaultUvsTable.Ranges)
                    {
                        // A default-UVS range can span many code points; only the ones actually used
                        // in this subset are kept, each re-emitted as its own single-value range
                        // (AdditionalCount = 0). This produces more, smaller ranges than the original
                        // font might use, but keeps the logic simple and correct - re-compacting
                        // adjacent surviving code points back into wider ranges isn't worth the
                        // complexity here.
                        uint rangeEnd = range.StartUnicodeValue + (uint)range.AdditionalCount;
                        for (uint cp = range.StartUnicodeValue; cp <= rangeEnd; cp++)
                        {
                            if (context.UsedCodePoints.Contains(cp))
                            {
                                if (newDefault == null)
                                    newDefault = new DefaultUvsTable { Ranges = new List<UnicodeRange>() };

                                newDefault.Ranges.Add(new UnicodeRange { StartUnicodeValue = cp, AdditionalCount = 0 });
                            }
                        }
                    }
                }

                if (newNonDefault == null && newDefault == null)
                    continue;

                newSubtable14.VariationSelectors.Add(new VariationSelector
                {
                    VarSelector = selector.VarSelector,
                    NonDefaultUvsTable = newNonDefault,
                    DefaultUvsTable = newDefault
                });
            }

            return newSubtable14.VariationSelectors.Count == 0 ? null : newSubtable14;
        }

        private CmapSubtable12 CreateFormat12Subtable(Dictionary<uint, ushort> mapping)
        {
            var subtable = new CmapSubtable12();

            // Sort by code point
            var sortedMappings = mapping.OrderBy(kvp => kvp.Key).ToList();

            if (sortedMappings.Count == 0)
            {
                subtable.NumGroups = 0;
                subtable.Length = 16; // Header only
                return subtable;
            }

            // Build sequential groups
            uint currentStart = sortedMappings[0].Key;
            uint currentStartGid = sortedMappings[0].Value;
            uint currentEnd = currentStart;

            for (int i = 1; i < sortedMappings.Count; i++)
            {
                uint codePoint = sortedMappings[i].Key;
                ushort glyphId = sortedMappings[i].Value;

                // Check if this continues the current sequential group
                bool isSequential = (codePoint == currentEnd + 1) &&
                                   (glyphId == currentStartGid + (codePoint - currentStart));

                if (isSequential)
                {
                    // Extend current group
                    currentEnd = codePoint;
                }
                else
                {
                    // Save current group and start new one
                    subtable.Groups.Add(new SequencialMapGroup
                    {
                        StartCharCode = currentStart,
                        EndCharCode = currentEnd,
                        StartGlyphId = currentStartGid
                    });

                    currentStart = codePoint;
                    currentStartGid = glyphId;
                    currentEnd = codePoint;
                }
            }

            // Add final group
            subtable.Groups.Add(new SequencialMapGroup
            {
                StartCharCode = currentStart,
                EndCharCode = currentEnd,
                StartGlyphId = currentStartGid
            });

            // Update metadata
            subtable.NumGroups = (uint)subtable.Groups.Count;

            // Calculate length: header (16 bytes) + groups (12 bytes each)
            subtable.Length = 16 + (subtable.NumGroups * 12);

            return subtable;
        }
    }
}