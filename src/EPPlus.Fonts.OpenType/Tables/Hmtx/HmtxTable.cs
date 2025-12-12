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
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Hmtx
{
    /// <summary>
    /// Glyph metrics used for horizontal text layout include glyph advance widths, side bearings 
    /// and X-direction min and max values (xMin, xMax). These are derived using a combination of 
    /// the glyph outline data ('glyf', 'CFF ' or CFF2) and the horizontal metrics table. The horizontal 
    /// metrics ('hmtx') table provides glyph advance widths and left side bearings.
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/hmtx
    /// </summary>
    public class HmtxTable : FontTableBase
    {
        public override string Name => TableNames.Hmtx;

        public override bool IsEssentialTable => false;
        public List<LongHorMetric> hMetrics { get; set; } = new List<LongHorMetric>();
        public List<short> leftSideBearings { get; set; } = new List<short>();

        internal override void Clear()
        {
            hMetrics.Clear();
            leftSideBearings.Clear();
        }

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            foreach (var metric in hMetrics)
            {
                writer.WriteUInt16BigEndian(metric.advanceWidth);
                writer.WriteInt16BigEndian(metric.lsb);
            }

            foreach (var lsb in leftSideBearings)
            {
                writer.WriteInt16BigEndian(lsb);
            }
        }


        public HmtxTable CloneSubset(HashSet<ushort> glyphSet, HmtxTable original)
        {
            var newTable = new HmtxTable();

            // Sort glyphs för stabilitet
            var sortedGlyphs = glyphSet.OrderBy(g => g).ToList();

            foreach (var glyphId in sortedGlyphs)
            {
                if (glyphId < original.hMetrics.Count)
                {
                    // Glyph har en full metric-post
                    var metric = original.hMetrics[glyphId];
                    newTable.hMetrics.Add(new LongHorMetric
                    {
                        advanceWidth = metric.advanceWidth,
                        lsb = metric.lsb
                    });
                }
                else
                {
                    // Glyph ligger utanför hMetrics -> hämta från leftSideBearings
                    int lsbIndex = glyphId - original.hMetrics.Count;
                    short lsbValue = (lsbIndex < original.leftSideBearings.Count)
                        ? original.leftSideBearings[lsbIndex]
                        : (short)0;

                    newTable.hMetrics.Add(new LongHorMetric
                    {
                        advanceWidth = original.hMetrics.Last().advanceWidth, // enligt spec
                        lsb = lsbValue
                    });
                }
            }

            return newTable;
        }

        internal HmtxTable CloneForGlyphCount(int newGlyphCount, int originalGlyphCount)
        {
            var clone = new HmtxTable();

            int originalHMetricsCount = this.hMetrics.Count;
            int originalLsbCount = this.leftSideBearings.Count;

            // hMetrics måste alltid ha newGlyphCount entries
            clone.hMetrics = new List<LongHorMetric>(newGlyphCount);

            // leftSideBearings får max newGlyphCount - originalHMetricsCount
            int neededLsbCount = Math.Max(0, newGlyphCount - originalHMetricsCount);
            clone.leftSideBearings = new List<short>(neededLsbCount);

            // Kopiera hMetrics
            int copyHMetrics = Math.Min(newGlyphCount, originalHMetricsCount);
            for (int i = 0; i < copyHMetrics; i++)
            {
                clone.hMetrics.Add(new LongHorMetric
                {
                    advanceWidth = this.hMetrics[i].advanceWidth,
                    lsb = this.hMetrics[i].lsb
                });
            }

            // Fyll på hMetrics med sista advanceWidth
            if (newGlyphCount > copyHMetrics)
            {
                var last = this.hMetrics[originalHMetricsCount - 1];
                for (int i = copyHMetrics; i < newGlyphCount; i++)
                {
                    clone.hMetrics.Add(new LongHorMetric
                    {
                        advanceWidth = last.advanceWidth,
                        lsb = 0
                    });
                }
            }

            // Kopiera befintliga LSB
            int copyLsb = Math.Min(neededLsbCount, originalLsbCount);
            for (int i = 0; i < copyLsb; i++)
            {
                clone.leftSideBearings.Add(this.leftSideBearings[i]);
            }

            // Fyll på med 0:or om vi har fler glyphs
            while (clone.leftSideBearings.Count < neededLsbCount)
            {
                clone.leftSideBearings.Add(0);
            }

            return clone;
        }

        public ushort GetAdvanceWidth(int glyphIndex)
        {
            if (glyphIndex < hMetrics.Count)
                return hMetrics[glyphIndex].advanceWidth;
            return hMetrics[hMetrics.Count - 1].advanceWidth; // fallback
        }

        public short GetLeftSideBearing(int glyphIndex)
        {
            if (glyphIndex < hMetrics.Count)
                return hMetrics[glyphIndex].lsb;
            if (glyphIndex < leftSideBearings.Count)
                return leftSideBearings[glyphIndex - hMetrics.Count];
            return 0;
        }
    }
}
