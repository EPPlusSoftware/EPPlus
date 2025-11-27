using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable12Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CmapSubtable12Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public CmapSubtable12 Deserialize(uint startIndex)
        {
            _reader.BaseStream.Position = startIndex;

            var subtable = new CmapSubtable12();

            // Read header
            var format = _reader.ReadUInt16BigEndian(); // should always be 12
            var reserved = _reader.ReadUInt16BigEndian(); // always 0
            subtable.Length = _reader.ReadUInt32BigEndian();
            subtable.Language = _reader.ReadUInt32BigEndian();
            subtable.NumGroups = _reader.ReadUInt32BigEndian();

            // Read groups
            for (int i = 0; i < subtable.NumGroups; i++)
            {
                var group = new SequencialMapGroup
                {
                    StartCharCode = _reader.ReadUInt32BigEndian(),
                    EndCharCode = _reader.ReadUInt32BigEndian(),
                    StartGlyphId = _reader.ReadUInt32BigEndian()
                };
                subtable.Groups.Add(group);
            }

            return subtable;
        }
    }
}
