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
namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    /// <summary>
    /// Each glyph description begins with a header
    /// </summary>
    public class GlyphHeader
    {
        internal GlyphHeader()
        {

        }

        internal GlyphHeader(short numberOfContours, BoundingRectangle rect)
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
