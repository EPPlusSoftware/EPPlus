/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution (Type 2) support
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class MultipleSubstSubTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public MultipleSubstSubTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public MultipleSubstSubTable Deserialize(long subTableStartOffset)
        {
            _reader.BaseStream.Seek(subTableStartOffset, SeekOrigin.Begin);
            long currentPos = subTableStartOffset;

            // USHORT SubstFormat
            ushort format = _reader.ReadUInt16BigEndian();
            if (format != 1)
            {
                throw new NotSupportedException($"Unsupported MultipleSubstSubTable format: {format}");
            }

            // USHORT CoverageOffset
            ushort coverageOffset = _reader.ReadUInt16BigEndian();

            // USHORT SequenceCount
            ushort sequenceCount = _reader.ReadUInt16BigEndian();

            // USHORT[] SequenceOffsets
            ushort[] sequenceOffsets = new ushort[sequenceCount];
            for (int i = 0; i < sequenceCount; i++)
            {
                sequenceOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            var subTable = new MultipleSubstSubTable
            {
                SubtableFormat = format,
                Sequences = new List<ushort[]>(sequenceCount)
            };

            // Each Sequence table is read at its OWN absolute offset - never by continuing to
            // read wherever the previous one left off, so there is nothing to restore between
            // iterations here.
            for (int i = 0; i < sequenceCount; i++)
            {
                long sequenceAbsoluteStart = subTableStartOffset + sequenceOffsets[i];
                _reader.BaseStream.Seek(sequenceAbsoluteStart, SeekOrigin.Begin);

                // USHORT GlyphCount
                ushort glyphCount = _reader.ReadUInt16BigEndian();

                // USHORT[] SubstituteGlyphIDs
                ushort[] sequence = new ushort[glyphCount];
                for (int g = 0; g < glyphCount; g++)
                {
                    sequence[g] = _reader.ReadUInt16BigEndian();
                }

                subTable.Sequences.Add(sequence);
            }

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