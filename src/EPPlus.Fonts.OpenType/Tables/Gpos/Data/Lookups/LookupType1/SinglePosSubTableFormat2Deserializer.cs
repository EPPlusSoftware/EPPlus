/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Lookup Type 1 Format 2 Deserializer
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1
{
    /// <summary>
    /// Deserializer for GPOS Lookup Type 1, Format 2: Single Adjustment Positioning
    /// Each glyph in coverage gets its own individual ValueRecord.
    /// </summary>
    internal class SinglePosSubTableFormat2Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public SinglePosSubTableFormat2Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        /// <summary>
        /// Deserializes a SinglePos Format 2 subtable from the current stream position.
        /// </summary>
        /// <param name="subtableStartOffset">Absolute offset where this subtable starts</param>
        /// <returns>Deserialized SinglePosSubTableFormat2</returns>
        public SinglePosSubTableFormat2 Deserialize(long subtableStartOffset)
        {
            // Seek to subtable start
            _reader.BaseStream.Seek(subtableStartOffset, SeekOrigin.Begin);

            var subtable = new SinglePosSubTableFormat2
            {
                SubtableFormat = _reader.ReadUInt16BigEndian(),
                CoverageOffset = _reader.ReadUInt16BigEndian(),
                ValueFormat = _reader.ReadUInt16BigEndian(),
                ValueCount = _reader.ReadUInt16BigEndian()
            };

            // Read array of ValueRecords
            subtable.Values = new ValueRecord[subtable.ValueCount];
            for (int i = 0; i < subtable.ValueCount; i++)
            {
                subtable.Values[i] = ReadValueRecord(subtable.ValueFormat);
            }

            // Read Coverage table
            if (subtable.CoverageOffset > 0)
            {
                long coveragePos = subtableStartOffset + subtable.CoverageOffset;
                _reader.BaseStream.Seek(coveragePos, SeekOrigin.Begin);

                ushort coverageFormat = _reader.ReadUInt16BigEndian();

                if (coverageFormat == 1)
                {
                    subtable.Coverage = new CoverageTableFormat1Deserializer(_reader)
                        .Deserialize(coveragePos);
                }
                else if (coverageFormat == 2)
                {
                    subtable.Coverage = new CoverageTableFormat2Deserializer(_reader)
                        .Deserialize(coveragePos);
                }
            }

            return subtable;
        }

        /// <summary>
        /// Reads a ValueRecord based on the ValueFormat flags.
        /// ValueFormat is a bit field indicating which fields are present.
        /// </summary>
        private ValueRecord ReadValueRecord(ushort valueFormat)
        {
            var record = new ValueRecord();

            // Bit 0x0001: XPlacement
            if ((valueFormat & 0x0001) != 0)
                record.XPlacement = _reader.ReadInt16BigEndian();

            // Bit 0x0002: YPlacement
            if ((valueFormat & 0x0002) != 0)
                record.YPlacement = _reader.ReadInt16BigEndian();

            // Bit 0x0004: XAdvance
            if ((valueFormat & 0x0004) != 0)
                record.XAdvance = _reader.ReadInt16BigEndian();

            // Bit 0x0008: YAdvance
            if ((valueFormat & 0x0008) != 0)
                record.YAdvance = _reader.ReadInt16BigEndian();

            // Bits 0x0010-0x0080: Device tables (not implemented yet)
            // Skip for now - we only need basic positioning

            return record;
        }
    }
}