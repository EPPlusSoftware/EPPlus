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
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting.Processors
{
    /// <summary>
    /// Creates a new cmap table (format 4) for the subset font subset.
    /// Uses only the code points actually needed + .notdef mapping.
    /// Must run after GlyfAndLocaSubsetProcessor (needs OldToNewGlyphId mapping).
    /// .NET 3.5 compatible.
    /// </summary>
    internal class CmapSubsetProcessor : IFontSubsetProcessor
    {
        public void Process(FontSubsettingContext context)
        {
            // Build mapping: Unicode code point → new glyph ID in subset
            Dictionary<uint, ushort> cmapMapping = new Dictionary<uint, ushort>();

            // Map all used code points
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
                else
                {
                    cmapMapping[codePoint] = 0; // .notdef fallback
                }
            }

            // Always map code point 0 to .notdef (required by spec)
            cmapMapping[0] = 0;

            // Create format 4 subtable – the only one we support (and the only one needed)
            CmapSubtable4 format4 = CmapFormat4.CreateFromMappings(cmapMapping);

            // Build new cmap table with both required encoding records
            CmapTable newCmap = new CmapTable();
            newCmap.Version = 0;
            newCmap.NumTables = 2; // Windows + Unicode

            // (3,1) – Windows Unicode BMP
            EncodingRecord winRecord = new EncodingRecord(Platforms.Windows, 1, 0);
            winRecord.Subtable = format4;

            // (0,3) – Unicode BMP (old Macintosh style, still used by many apps)
            EncodingRecord unicodeRecord = new EncodingRecord(Platforms.Unicode, 3, 0);
            unicodeRecord.Subtable = format4;

            newCmap.EncodingRecords.Add(winRecord);
            newCmap.EncodingRecords.Add(unicodeRecord);
            newCmap.SubTables.Add(format4);

            // Replace cmap in subset font
            context.SubsetFont.AddOrReplaceTable(newCmap);
        }
    }
}
