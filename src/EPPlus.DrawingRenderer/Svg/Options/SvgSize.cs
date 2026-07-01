/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using EPPlus.Fonts.OpenType.Utils;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgSize
    {
        /// <summary>
        /// Overrides the width for ouput the svg image.
        /// </summary>
        public double? Width { get; set; } = null;
        /// <summary>
        /// Overrides the height of the ouput svg drawing.
        /// </summary>
        public double? Height { get; set; } = null;
        public SvgSizeUnit Unit { get; set; } = SvgSizeUnit.Pixels;
        public double WidthPixels 
        {
            get
            {
                return GetPixels(Width ?? 0D, Unit);
            }
        }
        public double HeightPixels
        {
            get
            {
                return GetPixels(Height ?? 0D, Unit);
            }
        }

        internal static double GetPixels(double width, SvgSizeUnit unit)
        {
            switch(unit)
            {
                case SvgSizeUnit.Points:
                    return width.PointToPixel();
                case SvgSizeUnit.Inches:
                    return width * 96;
                case SvgSizeUnit.Centimeters:
                    return width * 96 / 2.54;
                case SvgSizeUnit.Millimeters:
                    return width * 96 / 25.4;
                default: //Pixels
                    return width;
            }
        }
    }
}

    public enum SvgSizeUnit
    {
        Pixels=0,
        Points=1,
        Millimeters,  // mm
        Centimeters,  // cm
        Inches        // in
}
