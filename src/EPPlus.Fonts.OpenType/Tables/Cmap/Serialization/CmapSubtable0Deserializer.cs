using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable0Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CmapSubtable0Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public CmapSubtable0 Deserialize(uint startIndex)
        {
            _reader.BaseStream.Seek(startIndex, SeekOrigin.Begin);

            var table = new CmapSubtable0
            {
                Length = _reader.ReadUInt16BigEndian(),
                Language = _reader.ReadUInt16BigEndian()
            };

            // Format 0 always has 256 bytes for glyphIdArray
            table.GlyphIdArray = _reader.ReadBytes(256);

            return table;
        }
    }
}
