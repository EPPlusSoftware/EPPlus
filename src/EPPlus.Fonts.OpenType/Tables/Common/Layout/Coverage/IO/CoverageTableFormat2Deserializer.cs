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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO
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
