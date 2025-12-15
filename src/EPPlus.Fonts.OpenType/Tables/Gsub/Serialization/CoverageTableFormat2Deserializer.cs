using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Serialization
{
    internal class CoverageTableFormat2Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CoverageTableFormat2Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public CoverageTableFormat2 Deserialize(long startIndex)
        {
            _reader.BaseStream.Seek(startIndex, SeekOrigin.Begin);

            // Read Format (already known to be 2, but read to advance position)
            ushort format = _reader.ReadUInt16BigEndian();

            CoverageTableFormat2 table = new CoverageTableFormat2 { CoverageFormat = format };

            // USHORT RangeCount
            table.RangeCount = _reader.ReadUInt16BigEndian();

            // Read RangeRecords
            for (int i = 0; i < table.RangeCount; i++)
            {
                CoverageRangeRecord record = new CoverageRangeRecord
                {
                    // USHORT StartGlyphID
                    StartGlyphID = _reader.ReadUInt16BigEndian(),
                    // USHORT EndGlyphID
                    EndGlyphID = _reader.ReadUInt16BigEndian(),
                    // USHORT StartCoverageIndex
                    StartCoverageIndex = _reader.ReadUInt16BigEndian()
                };
                table.RangeRecords.Add(record);
            }
            return table;
        }
    }
}
