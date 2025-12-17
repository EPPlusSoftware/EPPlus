using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class CoverageTableFormat1Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CoverageTableFormat1Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public CoverageTableFormat1 Deserialize(long startIndex)
        {
            _reader.BaseStream.Seek(startIndex, SeekOrigin.Begin);

            // Read Format (already known to be 1, but read to advance position)
            ushort format = _reader.ReadUInt16BigEndian();

            CoverageTableFormat1 table = new CoverageTableFormat1 { CoverageFormat = format };

            // USHORT GlyphCount
            table.GlyphCount = _reader.ReadUInt16BigEndian();

            // USHORT[] GlyphArray
            table.GlyphArray = new ushort[table.GlyphCount];

            for (int i = 0; i < table.GlyphCount; i++)
            {
                table.GlyphArray[i] = _reader.ReadUInt16BigEndian();
            }

            return table;
        }
    }
}
