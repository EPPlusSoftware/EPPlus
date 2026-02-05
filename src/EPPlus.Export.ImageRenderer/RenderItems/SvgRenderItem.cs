/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System.Globalization;
using System.Linq;
using System.Text;
namespace EPPlusImageRenderer.RenderItems
{
    internal enum SvgFillType
    {
        SolidFill,
        GradientFill,
        PatternFill
    }
    internal abstract class SvgRenderItem : RenderItem
    {
        internal SvgRenderItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {
        }
        public override void Render(StringBuilder sb)
        {
            if (string.IsNullOrEmpty(FillColor) == false)
            {
                sb.Append($"fill=\"{FillColor}\" ");
            }
            //If fill is null it may in e.g. Rect still get the color black which can have an opacity
            if (FillOpacity != null && FillOpacity != 1)
            {
                sb.Append($"opacity=\"{FillOpacity.Value.ToString(CultureInfo.InvariantCulture)}\" ");
            }
            if (string.IsNullOrEmpty(FilterName) == false)
            {
                sb.Append($"filter=\"{FilterName}\" ");
            }
           
            if (BorderWidth.HasValue)
            {
                if (string.IsNullOrEmpty(BorderColor) == false)
                {
                    sb.Append($"stroke=\"{BorderColor}\" ");
                }
                var v = BorderWidth.Value * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL;
                sb.Append($"stroke-width=\"{v.ToString(CultureInfo.InvariantCulture)}\" ");

                if (BorderDashArray != null)
                {
                    var BorderDashArrayStr = BorderDashArray.Select(x => 
                    x.ToString(CultureInfo.InvariantCulture)).ToArray();

                    sb.Append($"stroke-dasharray=\"" + $"{string.Join(",", BorderDashArrayStr)}\" ");
                }
            }

            sb.Append($"stroke-miterlimit =\"8\" ");
        }
        internal abstract SvgRenderItem Clone(SvgShape svgDocument);
    }
}