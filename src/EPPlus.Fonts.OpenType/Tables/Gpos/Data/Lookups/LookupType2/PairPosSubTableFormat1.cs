/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           GPOS PairPos Format 1
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2
{
    /// <summary>
    /// PairPos Format 1: Adjustments for glyph pairs listed explicitly.
    /// Used when there are few pairs (e.g., specific kerning pairs like AV, AW).
    /// </summary>
    public class PairPosSubTableFormat1 : PairPosSubTable
    {
        /// <summary>
        /// Pair sets indexed by coverage index
        /// </summary>
        public List<PairSet> PairSets { get; set; }

        public override bool TryGetPairAdjustment(
    ushort firstGlyph,
    ushort secondGlyph,
    out ValueRecord value1,
    out ValueRecord value2)
        {
            value1 = null;
            value2 = null;

            // Check if first glyph is covered
            int coverageIndex = Coverage?.GetGlyphIndex(firstGlyph) ?? -1;
            if (coverageIndex < 0 || coverageIndex >= (PairSets?.Count ?? 0))
                return false;

            var pairSet = PairSets[coverageIndex];
            if (pairSet?.PairValueRecords == null || pairSet.PairValueRecords.Count == 0)
                return false;

            // Binary search for matching second glyph
            // PairValueRecords are guaranteed sorted by SecondGlyph per OpenType spec
            int left = 0;
            int right = pairSet.PairValueRecords.Count - 1;

            while (left <= right)
            {
                int mid = left + (right - left) / 2;
                ushort midGlyphId = pairSet.PairValueRecords[mid].SecondGlyph;

                if (midGlyphId == secondGlyph)
                {
                    // Found match
                    var record = pairSet.PairValueRecords[mid];
                    value1 = record.Value1;
                    value2 = record.Value2;
                    return true;
                }

                if (midGlyphId < secondGlyph)
                    left = mid + 1;
                else
                    right = mid - 1;
            }

            return false;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long subtableStart = writer.BaseStream.Position;

            // Write header
            writer.WriteUInt16BigEndian(SubtableFormat); // 1

            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Coverage offset placeholder

            writer.WriteUInt16BigEndian(ValueFormat1);
            writer.WriteUInt16BigEndian(ValueFormat2);
            writer.WriteUInt16BigEndian((ushort)PairSets.Count);

            // Reserve space for PairSet offsets
            long pairSetOffsetsPos = writer.BaseStream.Position;
            for (int i = 0; i < PairSets.Count; i++)
            {
                writer.WriteUInt16BigEndian(0); // Placeholder
            }

            // Write PairSets
            for (int i = 0; i < PairSets.Count; i++)
            {
                var pairSet = PairSets[i];
                if (pairSet == null || pairSet.PairValueRecords.Count == 0)
                {
                    continue; // Leave offset as 0
                }

                long pairSetStart = writer.BaseStream.Position;
                ushort pairSetOffset = (ushort)(pairSetStart - subtableStart);

                // Write PairSet
                writer.WriteUInt16BigEndian((ushort)pairSet.PairValueRecords.Count);

                foreach (var record in pairSet.PairValueRecords)
                {
                    writer.WriteUInt16BigEndian(record.SecondGlyph);
                    ValueRecordSerializer.Serialize(writer, record.Value1, ValueFormat1);
                    ValueRecordSerializer.Serialize(writer, record.Value2, ValueFormat2);
                }

                // Update PairSet offset
                long currentPos = writer.BaseStream.Position;
                writer.BaseStream.Seek(pairSetOffsetsPos + (i * 2), System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(pairSetOffset);
                writer.BaseStream.Seek(currentPos, System.IO.SeekOrigin.Begin);
            }

            // Write Coverage table
            if (Coverage != null)
            {
                ushort coverageOffset = (ushort)(writer.BaseStream.Position - subtableStart);
                long resumePos = writer.BaseStream.Position;

                // Update offset
                writer.BaseStream.Seek(coverageOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(coverageOffset);
                writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

                // Serialize coverage
                Coverage.Serialize(writer);
            }
        }
    }
}