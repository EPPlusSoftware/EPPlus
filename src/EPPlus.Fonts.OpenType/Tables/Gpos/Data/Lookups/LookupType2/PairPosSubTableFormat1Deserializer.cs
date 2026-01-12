/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/09/2026         EPPlus Software AB           GPOS PairPos Format 1 deserializer
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2
{
    /// <summary>
    /// Deserializes PairPos Format 1 subtables (explicit glyph pairs with kerning)
    /// </summary>
    internal class PairPosSubTableFormat1Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public PairPosSubTableFormat1Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public PairPosSubTableFormat1 Deserialize(long subtableStart)
        {
            _reader.BaseStream.Seek(subtableStart, SeekOrigin.Begin);

            var table = new PairPosSubTableFormat1();

            // Read header
            table.SubtableFormat = _reader.ReadUInt16BigEndian(); // Should be 1
            ushort coverageOffset = _reader.ReadUInt16BigEndian();
            table.ValueFormat1 = _reader.ReadUInt16BigEndian();
            table.ValueFormat2 = _reader.ReadUInt16BigEndian();
            ushort pairSetCount = _reader.ReadUInt16BigEndian();

            // Read PairSet offsets
            var pairSetOffsets = new ushort[pairSetCount];
            for (int i = 0; i < pairSetCount; i++)
            {
                pairSetOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            // Read Coverage table
            if (coverageOffset > 0)
            {
                long coveragePos = subtableStart + coverageOffset;
                _reader.BaseStream.Seek(coveragePos, SeekOrigin.Begin);

                ushort coverageFormat = _reader.ReadUInt16BigEndian();

                if (coverageFormat == 1)
                {
                    table.Coverage = new CoverageTableFormat1Deserializer(_reader)
                        .Deserialize(coveragePos);
                }
                else if (coverageFormat == 2)
                {
                    table.Coverage = new CoverageTableFormat2Deserializer(_reader)
                        .Deserialize(coveragePos);
                }
            }

            // Read PairSets
            table.PairSets = new List<PairSet>();

            for (int i = 0; i < pairSetCount; i++)
            {
                if (pairSetOffsets[i] == 0)
                {
                    table.PairSets.Add(null);
                    continue;
                }

                long pairSetPos = subtableStart + pairSetOffsets[i];
                _reader.BaseStream.Seek(pairSetPos, SeekOrigin.Begin);

                var pairSet = ReadPairSet(table.ValueFormat1, table.ValueFormat2);
                table.PairSets.Add(pairSet);
            }

            return table;
        }

        private PairSet ReadPairSet(ushort valueFormat1, ushort valueFormat2)
        {
            var pairSet = new PairSet();

            ushort pairValueCount = _reader.ReadUInt16BigEndian();
            pairSet.PairValueRecords = new List<PairValueRecord>();

            for (int i = 0; i < pairValueCount; i++)
            {
                var record = new PairValueRecord();

                // Read second glyph ID
                record.SecondGlyph = _reader.ReadUInt16BigEndian();

                // Read Value1 (adjustments for first glyph)
                record.Value1 = ValueRecord.Read(_reader, valueFormat1);

                // Read Value2 (adjustments for second glyph)
                record.Value2 = ValueRecord.Read(_reader, valueFormat2);

                pairSet.PairValueRecords.Add(record);
            }

            return pairSet;
        }
    }
}