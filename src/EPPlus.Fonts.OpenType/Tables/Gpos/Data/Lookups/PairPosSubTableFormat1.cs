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

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups
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
            if (pairSet?.PairValueRecords == null)
                return false;

            // Find matching second glyph
            foreach (var record in pairSet.PairValueRecords)
            {
                if (record.SecondGlyph == secondGlyph)
                {
                    value1 = record.Value1;
                    value2 = record.Value2;
                    return true;
                }
            }

            return false;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long subtableStart = writer.BaseStream.Position;

            // Write format
            writer.WriteUInt16BigEndian(1);

            // Write coverage offset placeholder
            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // Write value formats
            writer.WriteUInt16BigEndian(ValueFormat1);
            writer.WriteUInt16BigEndian(ValueFormat2);

            // PairSetCount
            writer.WriteUInt16BigEndian((ushort)(PairSets?.Count ?? 0));

            // Write PairSet offset placeholders
            var pairSetOffsetPositions = new List<long>();
            for (int i = 0; i < (PairSets?.Count ?? 0); i++)
            {
                pairSetOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // Write Coverage
            if (Coverage != null)
            {
                WriteRelativeOffset(writer, subtableStart, coverageOffsetPos);
                Coverage.Serialize(writer);
            }

            // Write PairSets
            if (PairSets != null)
            {
                for (int i = 0; i < PairSets.Count; i++)
                {
                    WriteRelativeOffset(writer, subtableStart, pairSetOffsetPositions[i]);
                    PairSets[i].Serialize(writer, ValueFormat1, ValueFormat2);
                }
            }
        }
    }
}