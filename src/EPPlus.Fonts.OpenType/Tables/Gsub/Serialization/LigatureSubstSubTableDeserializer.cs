using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Serialization
{
    internal class LigatureSubstSubTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public LigatureSubstSubTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public LigatureSubstSubTable Deserialize(long subTableStartOffset)
        {
            _reader.BaseStream.Seek(subTableStartOffset, SeekOrigin.Begin);

            LigatureSubstSubTable subTable = new LigatureSubstSubTable();

            // USHORT SubtableFormat (should be 1 for Format 1)
            subTable.SubtableFormat = _reader.ReadUInt16BigEndian();

            // USHORT CoverageOffset (relative to subTableStartOffset)
            ushort coverageOffset = _reader.ReadUInt16BigEndian();

            // USHORT LigSetCount
            ushort ligSetCount = _reader.ReadUInt16BigEndian();

            // Read LigatureSetOffsets (USHORT array, relative to subTableStartOffset)
            ushort[] ligSetOffsets = new ushort[ligSetCount];
            for (int i = 0; i < ligSetCount; i++)
            {
                ligSetOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            // 1. Deserialize CoverageTable (New logic)
            if (coverageOffset > 0)
            {
                long coverageAbsoluteStart = subTableStartOffset + coverageOffset;
                _reader.BaseStream.Seek(coverageAbsoluteStart, SeekOrigin.Begin);

                // Peek at the CoverageFormat to select the correct Deserializer
                ushort coverageFormat = _reader.ReadUInt16BigEndian();
                _reader.BaseStream.Seek(coverageAbsoluteStart, SeekOrigin.Begin); // Rewind

                if (coverageFormat == 1)
                {
                    var covDeserializer = new CoverageTableFormat1Deserializer(_reader);
                    subTable.Coverage = covDeserializer.Deserialize(coverageAbsoluteStart);
                }
                else if (coverageFormat == 2)
                {
                    var covDeserializer = new CoverageTableFormat2Deserializer(_reader);
                    subTable.Coverage = covDeserializer.Deserialize(coverageAbsoluteStart);
                }
                // Handle unsupported formats if necessary
            }

            // 2. Deserialize LigatureSetTables
            long currentPosition = _reader.BaseStream.Position;
            ushort[] coveredGlyphs = subTable.Coverage?.CoveredGlyphs ?? new ushort[0]; // Use new ushort[0] for .NET 3.5 safety

            if (coveredGlyphs.Length != ligSetCount)
            {
                // Log or handle error: CoverageTable must match LigatureSetCount
            }

            var ligSetDeserializer = new LigatureSetTableDeserializer(_reader); // Assuming you create this in the next step

            for (int i = 0; i < ligSetCount; i++)
            {
                // Navigate to LigatureSetTable: SubTable Start + LigatureSet Offset
                long ligSetAbsoluteStart = subTableStartOffset + ligSetOffsets[i];

                LigatureSetTable ligSet = ligSetDeserializer.Deserialize(ligSetAbsoluteStart);

                // Map LigatureSetTable to the BaseGlyph ID using the index (i) from CoverageTable
                if (coveredGlyphs.Length > i)
                {
                    ushort baseGlyphID = coveredGlyphs[i];
                    // This assumes LigatureSubstSubTable has a Dictionary<ushort, LigatureSetTable> property named LigatureSets
                    subTable.LigatureSets.Add(baseGlyphID, ligSet);
                }
            }

            // Restore position to after the offset array (optional, but clean)
            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);

            return subTable;
        }
    }
}
