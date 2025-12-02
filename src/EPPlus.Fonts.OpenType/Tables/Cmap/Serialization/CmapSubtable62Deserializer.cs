using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
