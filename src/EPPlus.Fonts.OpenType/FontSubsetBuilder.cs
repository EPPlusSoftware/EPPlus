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
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Loca;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType
{
    public class FontSubsetBuilder
    {
        private readonly OpenTypeFont _original;

        public FontSubsetBuilder(OpenTypeFont original)
        {
            _original = original;
        }

        public OpenTypeFont BuildSubset(HashSet<ushort> glyphIds, IEnumerable<char> usedChars)
        {
            var subsetFont = new OpenTypeFont(_original.Format);

            subsetFont.AddOrReplaceTable(_original.HeadTable.Clone());
            subsetFont.AddOrReplaceTable(_original.MaxpTable.Clone());
            subsetFont.MaxpTable.numGlyphs = (ushort)glyphIds.Count;

            // Build glyf subset
            var glyfProcessor = new GlyphSubsetProcessor(_original.GlyfTable);
            var glyphSubsetResult = glyfProcessor.CreateSubset(glyphIds);
            subsetFont.AddOrReplaceTable(glyphSubsetResult.GlyfTable);


            var glyfSize = glyphSubsetResult.GlyfTable.GetLength(subsetFont);
            var indexToLocFormat = glyfSize < 65536
                ? HeadTable.IndexToLocFormats.Offset16
                : HeadTable.IndexToLocFormats.Offset32;

            // Update HeadTable
            subsetFont.HeadTable.IndexToLocFormat = indexToLocFormat;

            // Build Loca-table for the subset
            subsetFont.AddOrReplaceTable(
                LocaTable.CreateSubset(glyphSubsetResult.LocaOffsets, indexToLocFormat)
            );

            return subsetFont;
        }
    }
}
