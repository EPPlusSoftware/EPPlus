namespace FontLab1.Tables.Hmtx
{
    /// <summary>
    /// Glyph metrics used for horizontal text layout include glyph advance widths, side bearings 
    /// and X-direction min and max values (xMin, xMax). These are derived using a combination of 
    /// the glyph outline data ('glyf', 'CFF ' or CFF2) and the horizontal metrics table. The horizontal 
    /// metrics ('hmtx') table provides glyph advance widths and left side bearings.
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/hmtx
    /// </summary>
    internal class HmtxTable
    {
        /// <summary>
        /// Paired advance width and left side bearing values for each glyph. Records are indexed by glyph ID.
        /// </summary>
        public LongHorMetric[] hMetrics { get; set; }

        /// <summary>
        /// Left side bearings for glyph IDs greater than or equal to numberOfHMetrics.
        /// </summary>
        public short[] leftSideBearings { get; set; }
    }
}
