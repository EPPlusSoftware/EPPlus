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
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable6Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CmapSubtable6Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public CmapSubtable6 Deserialize(uint startIndex)
        {
            _reader.BaseStream.Seek(startIndex, SeekOrigin.Begin);

            var table = new CmapSubtable6
            {
                Length = _reader.ReadUInt16BigEndian(),
                Language = _reader.ReadUInt16BigEndian(),
                FirstCode = _reader.ReadUInt16BigEndian(),
                EntryCount = _reader.ReadUInt16BigEndian()
            };

            table.GlyphIdArray = _reader.ReadUInt16ArrayBigEndian(table.EntryCount);

            return table;
        }
    }
}
