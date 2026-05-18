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
using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Renderer;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Globalization;
using System.Linq;
using System.Text;
using tc = OfficeOpenXml.Utils.TypeConversion;
using System.Drawing;
using OfficeOpenXml.Drawing.Renderer.RenderItems.Fill;
namespace EPPlusImageRenderer.RenderItems
{
    internal enum SvgFillType
    {
        SolidFill,
        GradientFill,
        PatternFill
    }
    internal abstract class DrawingRenderItem : RenderItem
    {
        //Refrence string if this is part of a definition
        internal string DefId = null;
        ExcelDrawing _drawing;
        ExcelTheme _theme;
        internal DrawingRenderItem(ExcelDrawing drawing, BoundingBox parent) : base(parent)
        {
            _drawing = drawing;
            _theme = _drawing._drawings.Worksheet.Workbook.ThemeManager.GetOrCreateTheme();
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

        //internal abstract DrawingRenderItem Clone(SvgShape svgDocument);
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
            if (LineCap != LineCap.Flat)
            {
                sb.AppendFormat(" stroke-linecap=\"{0}\"", LineCap == LineCap.Round ? "round" : "square");
            }
            if (LineJoin != LineJoin.Miter)
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
        internal void SetPatternFill()
        {

        }

        internal virtual void SetDrawingPropertiesFill(ExcelDrawingFill fill, ExcelDrawingColorManager color)
        {
            switch (fill.Style)
            {

                case eFillStyle.PatternFill:
                    PatternFill = new DrawingRenderPatternFill(fill.PatternFill);
                    break;
                case eFillStyle.BlipFill:
                    BlipFill = new DrawingRenderBlipFill(fill.BlipFill);
                    break;
                default:
                    SetDrawingPropertiesFill((ExcelDrawingFillBasic)fill, color);
                    break;
            }
        }
        internal virtual void SetDrawingPropertiesFill(ExcelDrawingFillBasic fill, ExcelDrawingColorManager color)
        {
            double? opacity = null;
            switch (fill.Style)
            {
                case eFillStyle.NoFill:
                    if (fill.IsEmpty)
                    {
                        FillColor = GetFillColor(fill, color, FillColorSource, out opacity);
                    }
                    else
                    {
                        FillColor = "none";
                    }
                    break;
                case eFillStyle.SolidFill:
                    FillColor = GetFillColor(fill, color, FillColorSource, out opacity);
                    break;
                case eFillStyle.GradientFill:
                    GradientFill = new DrawingRenderGradientFill(_theme, fill.GradientFill);
                    FillColor = null;
                    break;
            }
            if (opacity.HasValue)
            {
                FillOpacity = opacity;
            }
        }
        internal virtual void SetDrawingPropertiesBorder(ExcelDrawingBorder border, ExcelChartStyleColorManager color, bool hasBorder, double defaultWidth = 1.5)
        {
            double? opacity = null;
            switch (border.Fill.Style)
            {
                case eFillStyle.NoFill:
                    if (border.Fill.IsEmpty)
                    {
                        BorderColor = GetFillColor(border.Fill, color, BorderColorSource, out opacity);
                    }
                    else
                    {
                        BorderColor = "none";
                    }
                    break;
                case eFillStyle.SolidFill:
                    BorderColor = GetFillColor(border.Fill, color, BorderColorSource, out opacity);
                    BorderGradientFill = null;
                    break;
                case eFillStyle.GradientFill:
                    BorderGradientFill = new RenderGradientFill();
                    BorderColor = null;
                    break;
            }

            if (opacity.HasValue)
            {
                BorderOpacity = opacity;
            }

            if (hasBorder && BorderColorSource != PathFillMode.None)
            {
                BorderWidth = border.Width == 0 ? defaultWidth : border.Width;
                if (border.LineStyle.HasValue && border.LineStyle != eLineStyle.Solid)
                {
                    BorderDashArray = GetDashArray(border);
                }
                if (border.CompoundLineStyle != eCompoundLineStyle.Single)
                {
                    CompoundLineStyle = (CompoundLineStyle)border.CompoundLineStyle;
                    //TODO:Add support double compound borders.
                }
            }
        }
        internal void SetDrawingPropertiesEffects(ExcelDrawingEffectStyle effect)
        {
            if (effect.HasGlow)
            {
                GlowRadius = effect.Glow.Radius;
                var gc = tc.ColorConverter.GetThemeColor(_theme, effect.Glow.Color);
                GlowColor = "#" + gc.ToArgb().ToString("x8").Substring(2);
            }
            if (effect.HasOuterShadow)
            {
                OuterShadowEffect.OuterShadowEffectColor = tc.ColorConverter.GetThemeColor(_theme, effect.OuterShadow.Color);
            }
        }

        private double[] GetDashArray(ExcelDrawingBorder border)
        {
            var lw = (int)Math.Round(border.Width * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL);
            switch (border.LineStyle)
            {
                case eLineStyle.Dot:
                    return new double[] { lw, 4 * lw };
                case eLineStyle.DashDot:
                    return new double[] { 4 * lw, 3 * lw, lw, 3 * lw };
                case eLineStyle.Dash:
                    return new double[] { 4 * lw, 3 * lw };
                case eLineStyle.LongDash:
                    return new double[] { 8 * lw, 3 * lw };
                case eLineStyle.LongDashDot:
                    return new double[] { 8 * lw, 3 * lw, lw, 3 * lw };
                case eLineStyle.LongDashDotDot:
                    return new double[] { 8 * lw, 3 * lw, lw, 3 * lw, lw, 3 * lw };
                case eLineStyle.SystemDash:
                    return new double[] { 3 * lw, lw };
                case eLineStyle.SystemDot:
                    return new double[] { lw, lw };
                case eLineStyle.SystemDashDot:
                    return new double[] { 3 * lw, lw, lw, lw };
                case eLineStyle.SystemDashDotDot:
                    return new double[] { 3 * lw, lw, lw, lw, lw, lw };
            }
            return null;
        }

        private string GetFillColor(ExcelDrawingFillBasic fill, ExcelDrawingColorManager styleFillColor, PathFillMode fillColorSource, out double? opacity)
        {
            opacity = null;
            if (fillColorSource == PathFillMode.None)
            {
                return "none";
            }

            Color fc;
            if (fill == null || fill.Style == eFillStyle.NoFill)
            {
                if (styleFillColor == null)
                {
                    fc = tc.ColorConverter.GetThemeColor(_theme.ColorScheme.Accent1);
                }
                else
                {
                    fc = tc.ColorConverter.GetThemeColor(_theme, styleFillColor);
                }
            }
            else if (fill.Style == eFillStyle.SolidFill)
            {
                fc = tc.ColorConverter.GetThemeColor(_theme, fill.SolidFill.Color);
            }
            else
            {
                return string.Empty;
            }

            fc = tc.ColorConverter.GetAdjustedColor(fillColorSource, fc);
            if (fc.A < 255 && fc != Color.Empty)
            {
                opacity = fc.A / 255D;
            }
            return "#" + fc.ToArgb().ToString("x8").Substring(2);
        }

        internal void GetOuterShadowColor(out string shadowColor, out double opacity)
        {
            if (OuterShadowEffect == null)
            {
                shadowColor = null;
                opacity = 0;

            }
            else
            {
                var tc = OuterShadowEffect.OuterShadowEffectColor;
                if (tc.A < 255 && tc != Color.Empty)
                {
                    opacity = tc.A / 255D;
                }
                else
                {
                    opacity = 1;
                }
                shadowColor = "#" + tc.ToArgb().ToString("x8").Substring(2);
            }
        }
    }

}
