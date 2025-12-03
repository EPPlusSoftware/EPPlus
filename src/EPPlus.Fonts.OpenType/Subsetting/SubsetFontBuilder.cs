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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class SubsetFontBuilder
    {

        public OpenTypeFont CreateSubset(OpenTypeFont originalFont, IEnumerable<int> unicodeChars)
        {
            var glyphSet = BuildGlyphSet(originalFont, unicodeChars); // HashSet<ushort>

            // 1) head
            var newFont = new OpenTypeFont(originalFont.Format);
            var headTable = originalFont.HeadTable.Clone();
            newFont.AddOrReplaceTable(headTable);


            // 2) name
            if (originalFont.NameTable != null)
            {
                var nameTable = originalFont.NameTable.Clone();
                newFont.AddOrReplaceTable(nameTable);
            }

            // 3) maxp
            if (originalFont.MaxpTable != null)
            {
                var maxpTable = originalFont.MaxpTable.Clone();
                maxpTable.numGlyphs = (ushort)glyphSet.Count; // glyphSet from BuildGlyphSet
                newFont.AddOrReplaceTable(maxpTable);
            }

            // 4) hhea

            //if (originalFont.HheaTable != null)
            //{
            //    var hheaTable = originalFont.HheaTable.Clone();
            //    hheaTable.numberOfHMetrics = (ushort)glyphSet.Count; // Temporärt, tills hmtx är klar
            //    newFont.AddOrReplaceTable(hheaTable);
            //}

            return newFont;

        }


        private HashSet<ushort> BuildGlyphSet(OpenTypeFont font, IEnumerable<int> unicodeChars)
        {
            var glyphIds = new HashSet<ushort>();

            foreach (var codePoint in unicodeChars)
            {
                if (font.CmapTable.TryGetGlyphId(codePoint, out ushort glyphId))
                {
                    glyphIds.Add(glyphId);
                }
            }

            glyphIds.Add(0); // Always include .notdef

            font.GlyfTable.ResolveCompositeGlyphs(glyphIds);

            return glyphIds;
        }
    }
}
