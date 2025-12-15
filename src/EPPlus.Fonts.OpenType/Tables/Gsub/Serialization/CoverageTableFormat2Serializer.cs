using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Serialization
{
    internal class CoverageTableFormat2Serializer
    {
        public void Serialize(CoverageTableFormat2 table, FontsBinaryWriter writer)
        {
            if (table.RangeRecords == null || table.RangeRecords.Count == 0)
            {
                return;
            }

            // Format 2 requires RangeRecords to be sorted by StartGlyphID
            List<CoverageRangeRecord> sortedRecords = table.RangeRecords.OrderBy(r => r.StartGlyphID).ToList();

            // USHORT CoverageFormat (2)
            writer.WriteUInt16BigEndian(2);

            // USHORT RangeCount
            writer.WriteUInt16BigEndian((ushort)sortedRecords.Count);

            // CoverageRangeRecord[]
            // StartCoverageIndex MÅSTE vara 0, 1, 2, ... för de nya subsettabellerna.
            ushort currentCoverageIndex = 0;
            foreach (var record in sortedRecords)
            {
                // USHORT StartGlyphID
                writer.WriteUInt16BigEndian(record.StartGlyphID);
                // USHORT EndGlyphID
                writer.WriteUInt16BigEndian(record.EndGlyphID);
                // USHORT StartCoverageIndex (måste räknas om för subsetet)
                writer.WriteUInt16BigEndian(currentCoverageIndex);

                // Räkna ut antalet glyfer i intervallet för nästa index
                currentCoverageIndex += (ushort)(record.EndGlyphID - record.StartGlyphID + 1);
            }
        }
    }
}
