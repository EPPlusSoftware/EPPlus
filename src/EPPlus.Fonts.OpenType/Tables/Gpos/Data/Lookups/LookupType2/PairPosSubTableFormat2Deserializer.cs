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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.ClassDef.IO;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2
{
    /// <summary>
    /// Deserializes PairPos Format 2 subtables (class-based kerning)
    /// </summary>
    internal class PairPosSubTableFormat2Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public PairPosSubTableFormat2Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public PairPosSubTableFormat2 Deserialize(long subtableStart)
        {
            _reader.BaseStream.Seek(subtableStart, SeekOrigin.Begin);

            var table = new PairPosSubTableFormat2();
            // Header 
            table.SubtableFormat = _reader.ReadUInt16BigEndian(); // Should be 2 

            ushort coverageOffset = _reader.ReadUInt16BigEndian();
            table.ValueFormat1 = _reader.ReadUInt16BigEndian();
            table.ValueFormat2 = _reader.ReadUInt16BigEndian();

            ushort classDef1Offset = _reader.ReadUInt16BigEndian();
            ushort classDef2Offset = _reader.ReadUInt16BigEndian();

            table.Class1Count = _reader.ReadUInt16BigEndian();
            table.Class2Count = _reader.ReadUInt16BigEndian();

            // Read class matrix
            table.ClassMatrix = new PairValueRecord[table.Class1Count, table.Class2Count];

            for (int i = 0; i < table.Class1Count; i++)
            {
                for (int j = 0; j < table.Class2Count; j++)
                {
                    var value1 = ValueRecord.Read(_reader, table.ValueFormat1);
                    var value2 = ValueRecord.Read(_reader, table.ValueFormat2);

                    table.ClassMatrix[i, j] = new PairValueRecord
                    {
                        Value1 = value1,
                        Value2 = value2
                    };
                }
            }

            // Coverage
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

            // ClassDef1
            if (classDef1Offset > 0)
            {
                long classDef1Pos = subtableStart + classDef1Offset;
                _reader.BaseStream.Seek(classDef1Pos, SeekOrigin.Begin);

                table.ClassDef1 = new ClassDefTableDeserializer(_reader)
                    .Deserialize(classDef1Pos);
            }

            // ClassDef2
            if (classDef2Offset > 0)
            {
                long classDef2Pos = subtableStart + classDef2Offset;
                _reader.BaseStream.Seek(classDef2Pos, SeekOrigin.Begin);

                table.ClassDef2 = new ClassDefTableDeserializer(_reader)
                    .Deserialize(classDef2Pos);
            }

            return table;
        }
    }
}
