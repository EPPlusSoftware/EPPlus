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
            // --- PHASE 3: REWRITE ---
            // Now context.OldToNewGlyphId IS populated. We can safely map 
            // Unicode -> OldGID -> NewGID.

            // Build mapping: Unicode code point → NEW glyph ID in subset
            Dictionary<uint, ushort> cmapMapping = new Dictionary<uint, ushort>();

            foreach (uint codePoint in context.UsedCodePoints)
            {
                ushort oldGid;
                if (context.OriginalFont.CmapTable.TryGetGlyphId(codePoint, out oldGid))
                {
                    ushort newGid;
                    // Map the old ID to the new dense ID (0, 1, 2...)
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

            // Create format 4 subtable using the provided class names
            // This creates the internal segment structure (Start/EndCount, IdDelta, etc.)
            CmapSubtable4 format4 = CmapFormat4.CreateFromMappings(cmapMapping);

            // Build new cmap table
            CmapTable newCmap = new CmapTable();
            newCmap.Version = 0;
            // Note: NumTables is usually updated automatically when records are added, 
            // but we set it to be explicit.
            newCmap.NumTables = 2;

            // (3,1) – Windows Unicode BMP
            EncodingRecord winRecord = new EncodingRecord(Platforms.Windows, 1, 0);
            winRecord.Subtable = format4;

            // (0,3) – Unicode BMP
            EncodingRecord unicodeRecord = new EncodingRecord(Platforms.Unicode, 3, 0);
            unicodeRecord.Subtable = format4;

            newCmap.EncodingRecords.Add(winRecord);
            newCmap.EncodingRecords.Add(unicodeRecord);

            // Add the subtable to the table's internal list
            newCmap.SubTables.Add(format4);

            // Replace cmap in subset font. 
            // Now GetMinCharCode() will return the correct value to the validator.
            context.SubsetFont.AddOrReplaceTable(newCmap);
        }
    }
}