using System.Diagnostics;

namespace FontLab1
{
    [DebuggerDisplay("x: ({Xmin} to {Xmax}), y: ({Ymin} to {Ymax})")]
    internal struct BoundingRectangle
    {
        public BoundingRectangle(short xMin, short yMin, short xMax, short yMax)
        {
            Xmin = xMin;
            Ymin = yMin;
            Xmax = xMax;
            Ymax = yMax;
        }
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

        public static BoundingRectangle Empty => new BoundingRectangle(-1, -1, -1, -1);
    }
}
