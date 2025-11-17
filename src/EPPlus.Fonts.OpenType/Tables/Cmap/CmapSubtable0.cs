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
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable0 : CmapSubtableBase
    {
        public CmapSubtable0()
        {
            GlyphIdArray = new byte[256];
        }

        public override ushort Format { get { return 0; } }

        public override uint Length { get; internal set; }

        public override uint Language { get; internal set; }

        /// <summary>
        /// Maps character codes 0–255 to glyph indices.
        /// </summary>
        public byte[] GlyphIdArray { get; internal set; }


        public override GlyphMappings GetGlyphMappings()
        {
            var mapping = new GlyphMappings();

            for (uint charCode = 0; charCode < 256; charCode++)
            {
                ushort glyphIndex = GlyphIdArray[charCode];
                if (glyphIndex != 0)
                {
                    mapping.AddMapping(charCode, glyphIndex);
                }
            }

            return mapping;
        }

        internal override int MapCodePointToGlyph(int codePoint)
        {
            if (codePoint < 0 || codePoint > 255)
                return -1; // Utanför intervallet

            return GlyphIdArray[codePoint];
        }


        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable02Serializer();
            serializer.Serialize(this, writer);
        }

        public void BuildFromMappings(List<CharGlyphMapping> mappings)
        {
            GlyphIdArray = new byte[256]; // default zeros

            foreach (CharGlyphMapping map in mappings)
            {
                if (map.CharCode < 256)
                {
                    GlyphIdArray[map.CharCode] = (byte)map.GlyphId;
                }
            }
        }
    }
}
