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
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
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
