using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class FeatureListTable : FontTableElement
    {
        public List<FeatureRecord> FeatureRecords { get; set; } = new List<FeatureRecord>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. USHORT FeatureCount
            writer.WriteUInt16BigEndian((ushort)this.FeatureRecords.Count);

            // Placeholder for USHORT[] FeatureRecords (FeatureTag + FeatureTableOffset)
            List<long> recordOffsetPositions = new List<long>();
            long featureListStartOffset = writer.BaseStream.Position - sizeof(ushort);

            foreach (var record in this.FeatureRecords)
            {
                // Write FeatureTag (4 bytes)
                writer.Write(record.FeatureTag.ToBytes());

                // Placeholder for FeatureTableOffset (2 bytes)
                recordOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- Skriv ut FeatureTables ---

            // Vi behöver en sorterad lista över FeatureRecords.

            int recordIndex = 0;
            foreach (var record in this.FeatureRecords)
            {
                long currentOffset = writer.BaseStream.Position;

                // Beräkna offset: FeatureTable start - FeatureList start
                ushort relativeFeatureTableOffset = (ushort)(currentOffset - featureListStartOffset);

                // 1. Gå tillbaka och fyll i offseten i FeatureRecord
                long recordOffsetPos = recordOffsetPositions[recordIndex];
                writer.BaseStream.Seek(recordOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeFeatureTableOffset);

                // 2. Återställ positionen och serialisera FeatureTable
                writer.BaseStream.Seek(currentOffset, SeekOrigin.Begin);

                // Serialisera FeatureTable
                // record.FeatureTable must implement Serialize
                record.FeatureTable.Serialize(writer);

                recordIndex++;
            }
        }
    }
}
