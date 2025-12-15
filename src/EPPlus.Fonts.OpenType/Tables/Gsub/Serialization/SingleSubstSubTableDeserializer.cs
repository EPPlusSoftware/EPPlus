using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Serialization
{
    internal class SingleSubstSubTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public SingleSubstSubTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public SingleSubstSubTable Deserialize(long subTableStartOffset)
        {
            _reader.BaseStream.Seek(subTableStartOffset, SeekOrigin.Begin);
            long currentPos = subTableStartOffset;

            // USHORT SubtableFormat
            ushort format = _reader.ReadUInt16BigEndian();

            // USHORT CoverageOffset
            ushort coverageOffset = _reader.ReadUInt16BigEndian();

            SingleSubstSubTable subTable;

            if (format == 1)
            {
                subTable = new SingleSubstSubTableFormat1();
                // SSHORT DeltaGlyphID
                ((SingleSubstSubTableFormat1)subTable).DeltaGlyphID = _reader.ReadInt16BigEndian();
            }
            else if (format == 2)
            {
                subTable = new SingleSubstSubTableFormat2();
                // USHORT GlyphCount
                ushort glyphCount = _reader.ReadUInt16BigEndian();
                ((SingleSubstSubTableFormat2)subTable).GlyphCount = glyphCount;

                // USHORT[] SubstituteGlyphIDs
                ((SingleSubstSubTableFormat2)subTable).SubstituteGlyphIDs = new ushort[glyphCount];
                for (int i = 0; i < glyphCount; i++)
                {
                    ((SingleSubstSubTableFormat2)subTable).SubstituteGlyphIDs[i] = _reader.ReadUInt16BigEndian();
                }
            }
            else
            {
                throw new NotSupportedException($"Unsupported SingleSubstSubTable format: {format}");
            }

            subTable.SubtableFormat = format;

            // Deserialize CoverageTable
            if (coverageOffset > 0)
            {
                long coverageAbsoluteStart = subTableStartOffset + coverageOffset;
                _reader.BaseStream.Seek(coverageAbsoluteStart, SeekOrigin.Begin);

                ushort coverageFormat = _reader.ReadUInt16BigEndian();
                _reader.BaseStream.Seek(coverageAbsoluteStart, SeekOrigin.Begin);

                if (coverageFormat == 1)
                {
                    subTable.Coverage = new CoverageTableFormat1Deserializer(_reader).Deserialize(coverageAbsoluteStart);
                }
                else if (coverageFormat == 2)
                {
                    subTable.Coverage = new CoverageTableFormat2Deserializer(_reader).Deserialize(coverageAbsoluteStart);
                }
            }

            _reader.BaseStream.Seek(currentPos, SeekOrigin.Begin);
            return subTable;
        }
    }
}
