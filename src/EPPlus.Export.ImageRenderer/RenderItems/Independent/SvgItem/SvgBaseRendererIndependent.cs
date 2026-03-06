using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Independent.SvgItem
{
    internal static class SvgBaseRendererIndependent
    {
        public static void BaseRender(StringBuilder sb, RenderItemIndependent item)
        {
            if (string.IsNullOrEmpty(item.FillColor) == false)
            {
                sb.Append($"fill=\"{item.FillColor}\" ");
                if (item.FillOpacity != null && item.FillOpacity != 1)
                {
                    sb.Append($"opacity=\"{item.FillOpacity.Value.ToString(CultureInfo.InvariantCulture)}\" ");
                }
            }
            if (string.IsNullOrEmpty(item.FilterName) == false)
            {
                sb.Append($"filter=\"{item.FilterName}\" ");
            }

            if (item.BorderWidth.HasValue)
            {
                if (string.IsNullOrEmpty(item.BorderColor) == false)
                {
                    sb.Append($"stroke=\"{item.BorderColor}\" ");
                }
                var v = item.BorderWidth.Value * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL;
                sb.Append($"stroke-width=\"{v.ToString(CultureInfo.InvariantCulture)}\" ");

                if (item.BorderDashArray != null)
                {
                    var BorderDashArrayStr = item.BorderDashArray.Select(x =>
                    x.ToString(CultureInfo.InvariantCulture)).ToArray();

                    sb.Append($"stroke-dasharray=\"" + $"{string.Join(",", BorderDashArrayStr)}\" ");
                }
            }
            sb.Append($"stroke-miterlimit =\"8\" ");
        }
    }
}
