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

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable6 : CmapSubtableBase
    {
        public override ushort Format => 6;

        public override uint Length { get; internal set; }

        public override uint Language { get; internal set; }

        public ushort FirstCode { get; internal set; }

        public ushort EntryCount { get; internal set; }

        public ushort[] GlyphIdArray { get; internal set; } = new ushort[0];

        public override GlyphMappings GetGlyphMappings()
        {
            var mapping = new GlyphMappings();

            for (int i = 0; i < EntryCount && i < GlyphIdArray.Length; i++)
            {
                uint charCode = (uint)(FirstCode + i);
                ushort glyphIndex = GlyphIdArray[i];

                mapping.AddMapping(charCode, glyphIndex);
            }

            return mapping;
        }


        internal override int MapCodePointToGlyph(int codePoint)
        {
            if (codePoint < FirstCode || codePoint >= FirstCode + EntryCount)
                return -1; // Not found

            int index = codePoint - FirstCode;
            if (index >= 0 && index < GlyphIdArray.Length)
            {
                return GlyphIdArray[index];
            }

            return -1; // Not found
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable6_2Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
