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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Theme;
using System;
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
    internal static class DrawingRenderItemExtentions
    {
        internal static void SetDrawingPropertiesFill(this RenderItem item, ExcelTheme theme, ExcelDrawingFill fill, ExcelDrawingColorManager color)
        {
            switch (fill.Style)
            {

                case eFillStyle.PatternFill:
                    item.PatternFill = new DrawingRenderPatternFill(theme, fill.PatternFill, item.FillColorSource);
                    break;
                case eFillStyle.BlipFill:
                    item.BlipFill = new DrawingRenderBlipFill(fill.BlipFill);
                    break;
                default:
                    SetDrawingPropertiesFillBasic(item, theme, fill, color);
                    break;
            }
        }
        internal static void SetDrawingPropertiesFillBasic(this RenderItem item, ExcelTheme theme, ExcelDrawingFillBasic fill, ExcelDrawingColorManager color)
        {
            double? opacity = null;
            switch (fill.Style)
            {
                case eFillStyle.NoFill:
                    if (fill.IsEmpty) //Removed for now. 
                    {
                        item.FillColor = GetFillColor(theme, fill, color, item.FillColorSource, out opacity);
                    }
                    else
                    {
                        item.FillColor = "none";
                    }
                    break;
                case eFillStyle.SolidFill:
                    item.FillColor = GetFillColor(theme, fill, color, item.FillColorSource, out opacity);
                    break;
                case eFillStyle.GradientFill:
                    item.GradientFill = new DrawingRenderGradientFill(theme, fill.GradientFill);
                    item.FillColor = null;
                    break;
            }
            if (opacity.HasValue)
            {
                item.FillOpacity = opacity;
            }
        }
        internal static void SetDrawingPropertiesBorder(this RenderItem item, ExcelTheme theme, ExcelDrawingBorder border, ExcelChartStyleColorManager color, bool hasBorder, double defaultWidth = 1.5)
        {
            double? opacity = null;
            if (border == null)
            {
                if (hasBorder)
                {
                    item.BorderColor = GetFillColor(theme, null, color, item.BorderColorSource, out opacity, theme.ColorScheme.Dark1);
                }
            }
            else
            {
                switch (border.Fill.Style)
                {
                    case eFillStyle.NoFill:
                        if (border.Fill.IsEmpty)
                        {
                            item.BorderColor = GetFillColor(theme, border.Fill, color, item.BorderColorSource, out opacity);
                        }
                        else
                        {
                            item.BorderColor = "none";
                        }
                        break;
                    case eFillStyle.SolidFill:
                        item.BorderColor = GetFillColor(theme, border.Fill, color, item.BorderColorSource, out opacity);
                        item.BorderGradientFill = null;
                        break;
                    case eFillStyle.GradientFill:
                        item.BorderGradientFill = new RenderGradientFill();
                        item.BorderColor = null;
                        break;
                }
            }

            if (opacity.HasValue)
            {
                item.BorderOpacity = opacity;
            }

            if (hasBorder && item.BorderColorSource != PathFillMode.None)
            {
                item.BorderWidth = (border?.Width??0D) == 0D ? defaultWidth : border.Width;
                if (border!=null && border.LineStyle.HasValue && border.LineStyle != eLineStyle.Solid)
                {
                    item.BorderDashArray = GetDashArray(border, item.BorderWidth.Value);
                }
                if (border != null && border.CompoundLineStyle != eCompoundLineStyle.Single)
                {
                    item.CompoundLineStyle = (CompoundLineStyle)border.CompoundLineStyle;
                    //TODO:Add support double compound borders.
                }
            }
        }
        internal static void SetDrawingPropertiesEffects(this RenderItem item, ExcelTheme theme, ExcelDrawingEffectStyle effect)
        {
            if (effect.HasGlow)
            {
                item.GlowRadius = effect.Glow.Radius;
                var gc = tc.ColorConverter.GetThemeColor(theme, effect.Glow.Color);
                item.GlowColor = "#" + gc.ToArgb().ToString("x8").Substring(2);
            }
            if (effect.HasOuterShadow)
            {
                item.OuterShadowEffect = new RenderShadowEffect();
                item.OuterShadowEffect.OuterShadowEffectColor = tc.ColorConverter.GetThemeColor(theme, effect.OuterShadow.Color);
                item.OuterShadowEffect.Direction = effect.OuterShadow.Direction;
                item.OuterShadowEffect.BlurRadius = effect.OuterShadow.BlurRadius;
                item.OuterShadowEffect.Distance = effect.OuterShadow.Distance;
            }
        }

        private static double[] GetDashArray(ExcelDrawingBorder border, double width)
        {
            var lw = (int)Math.Round(width * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL);
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

        private static string GetFillColor(ExcelTheme theme, ExcelDrawingFillBasic fill, ExcelDrawingColorManager styleFillColor, PathFillMode fillColorSource, out double? opacity,  ExcelDrawingThemeColorManager nullColor = null)
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
                    fc = tc.ColorConverter.GetThemeColor(nullColor ?? theme.ColorScheme.Accent1);
                }
                else
                {
                    fc = tc.ColorConverter.GetThemeColor(theme, styleFillColor);
                }
            }
            else if (fill.Style == eFillStyle.SolidFill)
            {
                fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill.Color);
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
    }

}
