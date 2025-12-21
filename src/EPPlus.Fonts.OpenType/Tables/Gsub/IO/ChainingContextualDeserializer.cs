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
using EPPlus.Fonts.OpenType.Tables.Common.Coverage;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    /// <summary>
    /// Deserializes Chaining Contextual Substitution subtables from the GSUB table.
    /// </summary>
    internal class ChainingContextualDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public ChainingContextualDeserializer(FontsBinaryReader reader) => _reader = reader;

        /// <summary>
        /// Deserializes a Format 3 (Coverage-based) Chaining Contextual subtable.
        /// </summary>
        /// <param name="absoluteStart">The absolute byte offset in the stream where the subtable starts.</param>
        public ChainingContextualSubstFormat3 Deserialize(long absoluteStart)
        {
            _reader.BaseStream.Seek(absoluteStart, SeekOrigin.Begin);
            var format = _reader.ReadUInt16BigEndian();

            // Currently, only Format 3 (Coverage-based context) is supported
            if (format != 3) return null;

            var table = new ChainingContextualSubstFormat3();

            // 1. Read Backtrack Coverage offsets
            var backtrackCount = _reader.ReadUInt16BigEndian();
            var backtrackOffsets = ReadOffsets(backtrackCount);

            // 2. Read Input Coverage offsets
            var inputCount = _reader.ReadUInt16BigEndian();
            var inputOffsets = ReadOffsets(inputCount);

            // 3. Read Lookahead Coverage offsets
            var lookaheadCount = _reader.ReadUInt16BigEndian();
            var lookaheadOffsets = ReadOffsets(lookaheadCount);

            // 4. Read Substitution Lookup Records
            var substCount = _reader.ReadUInt16BigEndian();
            for (int i = 0; i < substCount; i++)
            {
                table.SubstLookupRecords.Add(new SubstLookupRecord
                {
                    SequenceIndex = _reader.ReadUInt16BigEndian(),
                    LookupListIndex = _reader.ReadUInt16BigEndian()
                });
            }

            // 5. Load Coverage tables (performed last as offsets are relative to the subtable start)
            table.BacktrackCoverages = LoadCoverages(absoluteStart, backtrackOffsets);
            table.InputCoverages = LoadCoverages(absoluteStart, inputOffsets);
            table.LookaheadCoverages = LoadCoverages(absoluteStart, lookaheadOffsets);

            return table;
        }

        private ushort[] ReadOffsets(int count)
        {
            var offsets = new ushort[count];
            for (int i = 0; i < count; i++) offsets[i] = _reader.ReadUInt16BigEndian();
            return offsets;
        }

        private List<CoverageTable> LoadCoverages(long baseOffset, ushort[] offsets)
        {
            var list = new List<CoverageTable>();
            foreach (var offset in offsets)
            {
                long absolutePos = baseOffset + offset;
                _reader.BaseStream.Seek(absolutePos, SeekOrigin.Begin);

                // Read the first two bytes to determine the Coverage table format
                ushort format = _reader.ReadUInt16BigEndian();

                CoverageTable coverage;
                if (format == 1)
                {
                    var loader = new CoverageTableFormat1Deserializer(_reader);
                    coverage = loader.Deserialize(absolutePos);
                }
                else if (format == 2)
                {
                    var loader = new CoverageTableFormat2Deserializer(_reader);
                    coverage = loader.Deserialize(absolutePos);
                }
                else
                {
                    throw new NotSupportedException($"Coverage format {format} is not supported.");
                }

                list.Add(coverage);
            }
            return list;
        }
    }
}