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

            // Read format field and ensure we are at the expected subtable
            var format = _reader.ReadUInt16BigEndian();
            if (format != 6)
                throw new InvalidDataException($"Unexpected cmap subtable format: {format} (expected 6).");

            var table = new CmapSubtable6
            {
                Length = _reader.ReadUInt16BigEndian(),
                Language = _reader.ReadUInt16BigEndian(),
                FirstCode = _reader.ReadUInt16BigEndian(),
                EntryCount = _reader.ReadUInt16BigEndian()
            };

            if (table.EntryCount > 0)
            {
                // Read exactly EntryCount entries; if stream is truncated, let ReadUInt16ArrayBigEndian throw/end up as EndOfStreamException
                table.GlyphIdArray = _reader.ReadUInt16ArrayBigEndian(table.EntryCount) ?? new ushort[0];

                if (table.GlyphIdArray.Length != table.EntryCount)
                    throw new EndOfStreamException("Not enough data to read glyphIdArray for cmap format 6.");
            }
            else
            {
                table.GlyphIdArray = new ushort[0];
            }

            return table;
        }
    }
}