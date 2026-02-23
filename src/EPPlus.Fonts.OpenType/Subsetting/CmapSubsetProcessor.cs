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

            context.SubsetFont.AddOrReplaceTable(newCmap);
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