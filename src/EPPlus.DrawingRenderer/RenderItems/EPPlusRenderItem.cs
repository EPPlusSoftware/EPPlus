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
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
namespace EPPlusImageRenderer.RenderItems
{
    internal enum SvgFillType
    {
        SolidFill,
        GradientFill,
        PatternFill
    }
    internal abstract class EPPlusRenderItem : RenderItem
    {
        //Refrence string if this is part of a definition
        internal string DefId = null;

        internal EPPlusRenderItem(DrawingBase renderer, BoundingBox parent) : base(parent)
        {

        }
        public override void Render(StringBuilder sb)
        {
            RenderBase(sb);
        }

        private void RenderBase(StringBuilder sb)
        {
            if(Bounds.Name != null)
            {
                sb.Append($" id=\"{Bounds.Name}\" ");
            }

            if (string.IsNullOrEmpty(DefId) == false)
            {
                sb.Append($"id=\"{DefId}\" ");
            }

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
                if (BorderOpacity.HasValue)
                {
                    sb.Append($" stroke-opacity=\"{(Math.Round(BorderOpacity.Value * 100)).ToString(CultureInfo.InvariantCulture)}%\" ");
                }
            }

            if (TransformOrigin != null)
            {
                sb.Append($" transform-origin=\"{TransformOrigin.X.ToString(CultureInfo.InvariantCulture)} {TransformOrigin.Y.ToString(CultureInfo.InvariantCulture)}\" ");
            }

            sb.Append($"stroke-miterlimit =\"{StrokeMiterLimit}\" ");
        }

        internal abstract EPPlusRenderItem Clone(SvgShape svgDocument);
        private protected void RenderCompoundItems(StringBuilder sb, double? borderWidth, string color, string filter)
        {
            var tmpBorderWidth = BorderWidth;
            string tmpBorderColor = null;
            BorderWidth = borderWidth ?? BorderWidth;
            if (string.IsNullOrEmpty(color) == false)
            {
                tmpBorderColor = BorderColor;
                BorderColor = color;
            }

            RenderBase(sb);
            if (LineCap != eLineCap.Flat)
            {
                sb.AppendFormat(" stroke-linecap=\"{0}\"", LineCap == eLineCap.Round ? "round" : "square");
            }
            if (LineJoin != SvgLineJoin.Miter)
            {
                sb.AppendFormat(" stroke-linejoin=\"{0}\"", LineJoin);
            }

            if (string.IsNullOrEmpty(filter) == false)
            {
                sb.Append(" " + filter);
            }

            sb.AppendFormat("/>");

            BorderWidth = tmpBorderWidth;
            if (string.IsNullOrEmpty(color) == false)
            {
                BorderColor = tmpBorderColor;
            }
        }

    }
}