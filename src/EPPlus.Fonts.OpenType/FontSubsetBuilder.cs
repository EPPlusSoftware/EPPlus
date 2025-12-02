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
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Loca;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

            subsetFont.AddOrReplaceTable(TableNames.Head, _original.HeadTable.Clone());
            subsetFont.AddOrReplaceTable(TableNames.Maxp, _original.MaxpTable.Clone());
            subsetFont.MaxpTable.numGlyphs = (ushort)glyphIds.Count;

            // Build glyf subset
            var glyfProcessor = new GlyphSubsetProcessor(_original.GlyfTable);
            var glyphSubsetResult = glyfProcessor.CreateSubset(glyphIds);
            subsetFont.AddOrReplaceTable(TableNames.Glyf, glyphSubsetResult.GlyfTable);


            var glyfSize = glyphSubsetResult.GlyfTable.GetLength();
            var indexToLocFormat = glyfSize < 65536
                ? HeadTable.IndexToLocFormats.Offset16
                : HeadTable.IndexToLocFormats.Offset32;

            // Uppdatera HeadTable också
            subsetFont.HeadTable.IndexToLocFormat = indexToLocFormat;

            // Bygg Loca-tabellen
            subsetFont.AddOrReplaceTable(
                TableNames.Loca,
                LocaTable.CreateSubset(glyphSubsetResult.LocaOffsets, indexToLocFormat, subsetFont.MaxpTable)
            );

            //subsetFont.ReplaceTable(TableNames.Hmtx, BuildHmtxSubset(glyphIds));
            //subsetFont.ReplaceTable(TableNames.Cmap, BuildCmapSubset(usedChars));

            //subsetFont.RecalculateChecksums();
            return subsetFont;
        }
    }
}
