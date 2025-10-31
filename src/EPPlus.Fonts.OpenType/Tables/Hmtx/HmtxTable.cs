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
        /// <summary>
        /// Paired advance width and left side bearing values for each glyph. Records are indexed by glyph ID.
        /// </summary>
        public LongHorMetric[] hMetrics { get; set; }

        /// <summary>
        /// Left side bearings for glyph IDs greater than or equal to numberOfHMetrics.
        /// </summary>
        public short[] leftSideBearings { get; set; }

        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            // Write hMetrics: advanceWidth + lsb for each glyph
            foreach (var metric in hMetrics)
            {
                writer.WriteUInt16BigEndian(metric.advanceWidth);
                writer.WriteInt16BigEndian(metric.lsb);
            }

            // Write leftSideBearings: only lsb values for glyphs beyond numberOfHMetrics
            foreach (var lsb in leftSideBearings)
            {
                writer.WriteInt16BigEndian(lsb);
            }
        }
    }
}
