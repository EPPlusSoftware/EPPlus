using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class ChainingContextualDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public ChainingContextualDeserializer(FontsBinaryReader reader) => _reader = reader;

        public ChainingContextualSubstFormat3 Deserialize(long absoluteStart)
        {
            _reader.BaseStream.Seek(absoluteStart, SeekOrigin.Begin);
            var format = _reader.ReadUInt16BigEndian();
            if (format != 3) return null; // Vi börjar med Format 3

            var table = new ChainingContextualSubstFormat3();

            // 1. Läs Backtrack Coverage
            var backtrackCount = _reader.ReadUInt16BigEndian();
            var backtrackOffsets = ReadOffsets(backtrackCount);

            // 2. Läs Input Coverage
            var inputCount = _reader.ReadUInt16BigEndian();
            var inputOffsets = ReadOffsets(inputCount);

            // 3. Läs Lookahead Coverage
            var lookaheadCount = _reader.ReadUInt16BigEndian();
            var lookaheadOffsets = ReadOffsets(lookaheadCount);

            // 4. Läs Substitution Records
            var substCount = _reader.ReadUInt16BigEndian();
            for (int i = 0; i < substCount; i++)
            {
                table.SubstLookupRecords.Add(new SubstLookupRecord
                {
                    SequenceIndex = _reader.ReadUInt16BigEndian(),
                    LookupListIndex = _reader.ReadUInt16BigEndian()
                });
            }

            // 5. Fyll Coverage-tabellerna (Görs sist pga offsets är relativa till absoluteStart)
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

                // Läs de första två byten för att avgöra formatet
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
