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
namespace EPPlus.DrawingRenderer.Svg
{
    /// <summary>
    /// Options for rendering drawings to svg.
    /// </summary>
    public class SvgRenderOptions
    {
        /// <summary>
        /// The width of the drawing in pixels used for calculating output image. Overrides the width of the drawing if set. If not set, the width will be calculated based on the drawing.
        /// For output sizing, use the <see cref="SvgSize"/> property instead.
        /// </summary>
        public int? Width { get; set; }
        /// <summary>
        /// The height of the drawing in pixels. Overrides the height of the drawing if set. If not set, the width will be calculated based on the drawing.
        /// For output sizing, use the <see cref="SvgSize"/> property instead.
        /// </summary>
        public int? Height { get; set; }
        /// <summary>
        /// Sets the width and height of the svg image. If not set, the size will be calculated based on the drawings dimensions.
        /// </summary>
        public SvgSize SvgSize { get; } = new SvgSize();
    }
    
}