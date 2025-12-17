using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class LigatureSetTableDeserializer
    {
        private readonly FontsBinaryReader _reader;
        private readonly LigatureTableDeserializer _ligatureTableDeserializer;

        public LigatureSetTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
            _ligatureTableDeserializer = new LigatureTableDeserializer(reader);
        }

        public LigatureSetTable Deserialize(long ligSetStartOffset)
        {
            _reader.BaseStream.Seek(ligSetStartOffset, SeekOrigin.Begin);

            LigatureSetTable ligSet = new LigatureSetTable();

            // USHORT LigatureCount
            ushort ligCount = _reader.ReadUInt16BigEndian();

            // USHORT[] LigatureOffsets (relative to LigatureSetTable start)
            ushort[] ligOffsets = new ushort[ligCount];
            for (int i = 0; i < ligCount; i++)
            {
                ligOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            // Save current position after reading offsets
            long currentPosition = _reader.BaseStream.Position;

            // Deserialize all LigatureTables
            foreach (ushort offset in ligOffsets)
            {
                // Navigate to LigatureTable: LigatureSetTable Start + Ligature Offset
                long ligTableAbsoluteStart = ligSetStartOffset + offset;

                LigatureTable ligTable = _ligatureTableDeserializer.Deserialize(ligTableAbsoluteStart);
                ligSet.Ligatures.Add(ligTable);
            }

            // Restore position
            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);

            return ligSet;
        }
    }
}
