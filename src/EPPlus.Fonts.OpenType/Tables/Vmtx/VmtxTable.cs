
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           vmtx table implementation (vertical text support)
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Vmtx
{
    /// <summary>
    /// Represents the 'vmtx' (Vertical Metrics) table.
    /// Contains advance heights and top side bearings for each glyph.
    /// Only present in fonts with vertical layout support (primarily CJK).
    /// Analogous to the 'hmtx' table for horizontal metrics.
    /// </summary>
    public class VmtxTable : FontTableBase
    {
        public VmtxTable() { }

        public VmtxTable(List<LongVerMetric> vMetrics)
        {
            VMetrics = vMetrics;
            TopSideBearings = new List<short>();
        }

        public VmtxTable(List<LongVerMetric> vMetrics, List<short> topSideBearings)
        {
            VMetrics = vMetrics;
            TopSideBearings = topSideBearings;
        }

        /// <summary>
        /// Array of longVerMetric records, one per glyph up to numberOfVMetrics.
        /// Each record contains advanceHeight and topSideBearing.
        /// </summary>
        public List<LongVerMetric> VMetrics { get; set; } = new List<LongVerMetric>();

        /// <summary>
        /// Additional top side bearings for glyphs beyond numberOfVMetrics.
        /// These glyphs share the last advanceHeight in VMetrics.
        /// Count = maxp.numGlyphs - vhea.numberOfVMetrics.
        /// </summary>
        public List<short> TopSideBearings { get; set; } = new List<short>();

        public override string Name => TableNames.Vmtx;
        public override bool IsEssentialTable => false;

        /// <summary>
        /// Gets the advance height for a given glyph ID.
        /// If the glyph ID is beyond the VMetrics array, the last entry's advanceHeight is used
        /// (per OpenType spec).
        /// </summary>
        public ushort GetAdvanceHeight(ushort glyphId)
        {
            if (VMetrics == null || VMetrics.Count == 0)
                return 0;

            if (glyphId < VMetrics.Count)
                return VMetrics[glyphId].AdvanceHeight;

            // Glyphs beyond numberOfVMetrics share the last advanceHeight
            return VMetrics[VMetrics.Count - 1].AdvanceHeight;
        }

        /// <summary>
        /// Gets the top side bearing for a given glyph ID.
        /// Checks VMetrics first, then the TopSideBearings array.
        /// </summary>
        public short GetTopSideBearing(ushort glyphId)
        {
            if (VMetrics == null || VMetrics.Count == 0)
                return 0;

            if (glyphId < VMetrics.Count)
                return VMetrics[glyphId].TopSideBearing;

            // Look in the additional TopSideBearings array
            int tsb_idx = glyphId - VMetrics.Count;
            if (TopSideBearings != null && tsb_idx < TopSideBearings.Count)
                return TopSideBearings[tsb_idx];

            return 0;
        }

        /// <summary>
        /// Creates a deep copy of this VmtxTable.
        /// Used during font subsetting.
        /// </summary>
        public VmtxTable Clone()
        {
            var clonedMetrics = new List<LongVerMetric>(VMetrics.Count);
            foreach (var m in VMetrics)
                clonedMetrics.Add(new LongVerMetric { AdvanceHeight = m.AdvanceHeight, TopSideBearing = m.TopSideBearing });

            var clonedTsb = new List<short>(TopSideBearings);
            return new VmtxTable(clonedMetrics, clonedTsb);
        }

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            foreach (var m in VMetrics)
            {
                writer.WriteUInt16BigEndian(m.AdvanceHeight);
                writer.WriteInt16BigEndian(m.TopSideBearing);
            }
            foreach (var tsb in TopSideBearings)
            {
                writer.WriteInt16BigEndian(tsb);
            }
        }

        internal override void Clear()
        {
            // Not used in current architecture
        }
    }
}
