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
using EPPlus.Fonts.OpenType.Tables.Cmap.Serializers;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class CmapSubtable0 : CmapSubtableBase
    {
        internal CmapSubtable0(FontsBinaryReader reader)
        {
            _reader = reader;
            Format = 0;
            Length = _reader.ReadUInt16BigEndian();
            Language = _reader.ReadUInt16BigEndian();
            var mappings = new List<GlyphMapping>();
            for(var c = 0; c < 256; c++)
            {
                var b = reader.ReadByte();
                var ix = BitConverter.ToUInt16(new byte[] { b, 0 }, 0);
                if(ix != 0)
                {
                    mappings.Add(new GlyphMapping
                    {
                        CharacterCode = Convert.ToChar(c),
                        GlyphIndex = ix
                    });
                }
            }
            GlyphMappingArray = mappings.ToArray();
        }

        private readonly FontsBinaryReader _reader;

        public override ushort Format { get; }

        public override ushort Length { get; }

        public override ushort Language { get; }

        public override GlyphMapping[] GlyphMappingArray { get; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable0Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
