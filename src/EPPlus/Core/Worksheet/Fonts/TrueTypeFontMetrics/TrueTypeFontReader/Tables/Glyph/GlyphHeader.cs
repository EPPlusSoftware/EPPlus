namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Glyph
{
    /// <summary>
    /// Each glyph description begins with a header
    /// </summary>
    internal class GlyphHeader
    {
        public GlyphHeader()
        {

        }

        public GlyphHeader(short numberOfContours, BoundingRectangle rect)
        {
            this.numberOfContours = numberOfContours;
            xMin = rect.Xmin;
            xMax = rect.Xmax;
            yMin = rect.Ymin;
            yMax = rect.Ymax;
        }
        public short numberOfContours { get; set; }

        /// <summary>
        /// Minimum x for coordinate data.
        /// </summary>
        public short xMin { get; set; }

        /// <summary>
        /// Minimum y for coordinate data.
        /// </summary>
        public short yMin { get; set; }

        /// <summary>
        /// Maximum x for coordinate data.
        /// </summary>
        public short xMax { get; set; }

        /// <summary>
        /// Maximum y for coordinate data.
        /// </summary>
        public short yMax { get; set; }
    }
}
