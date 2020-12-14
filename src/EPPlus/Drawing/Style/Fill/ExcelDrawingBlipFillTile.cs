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
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// A BLIP will be tiled to fill the available space
    /// </summary>
    public class ExcelDrawingBlipFillTile
    {
        internal ExcelDrawingBlipFillTile()
        {

        }
        internal ExcelDrawingBlipFillTile(XmlHelper xml)
        {
            var v = xml.GetXmlNodeString("a:tile/@algn");
            if(!string.IsNullOrEmpty(v))
            {
                Alignment = v.TranslateRectangleAlignment();
            }
            else
            {
                Alignment = null;
            }
            FlipMode =  xml.GetXmlNodeString("a:tile/@flip").ToEnum(eTileFlipMode.None);
            HorizontalRatio = xml.GetXmlNodePercentage("a:tile/@sx") ?? 0;
            VerticalRatio = xml.GetXmlNodePercentage("a:tile/@sy") ?? 0;
            HorizontalOffset = (xml.GetXmlNodeDoubleNull("a:tile/@tx") ?? 0) / ExcelDrawing.EMU_PER_PIXEL;
            VerticalOffset = (xml.GetXmlNodeDoubleNull("a:tile/@ty") ?? 0) / ExcelDrawing.EMU_PER_PIXEL;
        }

        /// <summary>
        /// The direction(s) in which to flip the image.
        /// </summary>
        public eTileFlipMode? FlipMode { get; set; }
        /// <summary>
        /// Where to align the first tile with respect to the shape.
        /// </summary>
        public eRectangleAlignment? Alignment { get; set; }
        /// <summary>
        /// The ratio for horizontally scale
        /// </summary>
        public double HorizontalRatio { get; set; }
        /// <summary>
        /// The ratio for vertically scale
        /// </summary>
        public double VerticalRatio { get; set; }
        /// <summary>
        /// The horizontal offset after alignment
        /// </summary>
        public double HorizontalOffset { get; set; }
        /// <summary>
        /// The vertical offset after alignment
        /// </summary>
        public double VerticalOffset { get; set; }
    }
}