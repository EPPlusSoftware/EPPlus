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
        public List<LongHorMetric> hMetrics { get; set; } = new List<LongHorMetric>();
        public List<short> leftSideBearings { get; set; } = new List<short>();

        internal override void Clear()
        {
            hMetrics.Clear();
            leftSideBearings.Clear();
        }

        internal override void SerializeInternal(FontsBinaryWriter writer)
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

        public HmtxTable CreateSubset(HashSet<ushort> glyphIds, Dictionary<ushort, ushort> idMapping)
        {
            var sortedGlyphIds = glyphIds.OrderBy(id => id).ToList();
            var newHmtx = new HmtxTable();

            // Copy metrics for each glyph in subset
            foreach (var oldId in sortedGlyphIds)
            {
                LongHorMetric metric;
                if (oldId < hMetrics.Count)
                {
                    metric = new LongHorMetric
                    {
                        advanceWidth = hMetrics[oldId].advanceWidth,
                        lsb = hMetrics[oldId].lsb
                    };
                }
                else
                {
                    // Fallback: use .notdef metrics
                    metric = new LongHorMetric
                    {
                        advanceWidth = hMetrics[0].advanceWidth,
                        lsb = hMetrics[0].lsb
                    };
                }

                newHmtx.hMetrics.Add(metric);
            }

            // leftSideBearings are usually for glyphs beyond numOfLongHorMetrics
            // In subset, we can skip or keep empty because we only include used glyphs
            newHmtx.leftSideBearings.Clear();

            return newHmtx;
        }
    }
}
