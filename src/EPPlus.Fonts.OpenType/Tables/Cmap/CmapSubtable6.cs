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
using System.Linq;
using System;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class CmapSubtable6 : CmapSubtableBase
    {
        internal CmapSubtable6(FontsBinaryReader reader)
        {
            _reader = reader;
            Format = 6;
            // length is calculated so we just read an throw away...
            _reader.ReadUInt16BigEndian();
            Language = _reader.ReadUInt16BigEndian();
            var firstCode = _reader.ReadUInt16BigEndian();
            var entryCount = _reader.ReadUInt16BigEndian();
            //GlyphMappingArray = new GlyphMapping[entryCount];
            for(var x = 0; x < entryCount; x++)
            {
                //GlyphMappingArray[x] = new GlyphMapping
                //{
                //    CharacterCode = (char)(firstCode + x),
                //    GlyphIndex = _reader.ReadUInt16BigEndian()
                //};
            }
        }

        private readonly FontsBinaryReader _reader;

        public override ushort Format { get; }

        public override ushort Length { get; internal set; }// = (ushort)(10 + GlyphMappingArray.Length * 2);
        public override ushort Language { get; internal set; }

       // public override GlyphMapping[] GlyphMappingArray { get; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable6Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
