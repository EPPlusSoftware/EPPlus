namespace FontLab1.Tables.Head
{
    /// <summary>
    /// This table gives global information about the font.
    /// </summary>
    internal class HeadTable
    {
        public enum IndexToLocFormats : short
        {
            Offset16 = 0,
            Offset32 = 1
        }
        public ushort MajorVersion { get; set; }

        public ushort MinorVersion { get; set; }

        /// <summary>
        /// Set to a value from 16 to 16384. Any value in this range is valid. In fonts that have TrueType outlines, a power of 2 is recommended as this allows performance optimizations in some rasterizers.
        /// </summary>
        public ushort UnitsPerEm { get; set; }

        /// <summary>
        /// Minimum x coordinate across all glyph bounding boxes.
        /// </summary>
        public short Xmin { get; set; }

        /// <summary>
        /// Minimum y coordinate across all glyph bounding boxes.
        /// </summary>
        public short Ymin { get; set; }

        /// <summary>
        /// Maximum x coordinate across all glyph bounding boxes.
        /// </summary>
        public short Xmax { get; set; }

        /// <summary>
        /// Maximum y coordinate across all glyph bounding boxes.
        /// </summary>
        public short Ymax { get; set; }

        /// <summary>
        /// Smallest readable size in pixels.
        /// </summary>
        public ushort LowestRecPPEM { get; set; }

        /// <summary>
        /// 0 for short offsets (Offset16), 1 for long (Offset32).
        /// </summary>
        public IndexToLocFormats IndexToLocFormat { get; set; }

        public BoundingRectangle GetDefaultBounds()
        {
            return new BoundingRectangle(Xmin, Ymin, Xmax, Ymax);
        }
    }
}
