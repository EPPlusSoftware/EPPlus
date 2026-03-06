//using EPPlusImageRenderer.RenderItems;
//using EPPlusImageRenderer.Svg;
//using OfficeOpenXml.Drawing;
//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.Linq;
//using System.Text;

//namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
//{
//    internal abstract class SvgRenderItemIndependent : RenderItemIndependent
//    {
//        internal enum SvgFillType
//        {
//            SolidFill,
//            GradientFill,
//            PatternFill
//        }

//        public override void Render(StringBuilder sb)
//        {
//            if (string.IsNullOrEmpty(FillColor) == false)
//            {
//                sb.Append($"fill=\"{FillColor}\" ");
//            }
//            //If fill is null it may in e.g. Rect still get the color black which can have an opacity
//            if (FillOpacity != null && FillOpacity != 1)
//            {
//                sb.Append($"opacity=\"{FillOpacity.Value.ToString(CultureInfo.InvariantCulture)}\" ");
//            }
//            if (string.IsNullOrEmpty(FilterName) == false)
//            {
//                sb.Append($"filter=\"{FilterName}\" ");
//            }

//            if (BorderWidth.HasValue)
//            {
//                if (string.IsNullOrEmpty(BorderColor) == false)
//                {
//                    sb.Append($"stroke=\"{BorderColor}\" ");
//                }
//                var v = BorderWidth.Value * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL;
//                sb.Append($"stroke-width=\"{v.ToString(CultureInfo.InvariantCulture)}\" ");

//                if (BorderDashArray != null)
//                {
//                    var BorderDashArrayStr = BorderDashArray.Select(x =>
//                    x.ToString(CultureInfo.InvariantCulture)).ToArray();

//                    sb.Append($"stroke-dasharray=\"" + $"{string.Join(",", BorderDashArrayStr)}\" ");
//                }
//            }

//            sb.Append($"stroke-miterlimit =\"8\" ");
//        }
//        internal abstract SvgRenderItemIndependent Clone();
//    }
//}
