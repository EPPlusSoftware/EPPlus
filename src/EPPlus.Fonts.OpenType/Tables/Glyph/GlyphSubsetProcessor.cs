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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    internal class GlyphSubsetProcessor
    {
        private readonly GlyfTable _originalGlyphTable;

        public GlyphSubsetProcessor(GlyfTable originalGlyphTable)
        {
            _originalGlyphTable = originalGlyphTable;
        }

        public GlyphSubsetResult CreateSubset(HashSet<ushort> glyphIds)
        {
            var sortedIds = glyphIds.OrderBy(id => id).ToList();
            var newGlyphs = new List<Glyph>();
            var offsets = new List<uint>();

            uint currentOffset = 0;
            foreach (var id in sortedIds)
            {
                var glyph = _originalGlyphTable.GetGlyph(id);
                newGlyphs.Add(glyph);

                offsets.Add(currentOffset);
                currentOffset += AlignTo4Bytes(glyph.GetSize());
            }

            // Loca kräver en extra offset efter sista glyfen
            offsets.Add(currentOffset);

            var newGlyphTable = new GlyfTable(newGlyphs);
            return new GlyphSubsetResult
            {
                GlyfTable = newGlyphTable,
                LocaOffsets = offsets
            };
        }

        private uint AlignTo4Bytes(int size)
        {
            return (uint)((size + 3) & ~3);
        }
    }
}
