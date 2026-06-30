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
        internal static void SetDrawingPropertiesFill(this RenderItem item, ExcelTheme theme, ExcelDrawingFill fill, ExcelDrawingColorManager color, bool gradientUserSpace = false, Color? nullColor = null)
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
                    SetDrawingPropertiesFillBasic(item, theme, fill, color, gradientUserSpace, nullColor);
                    break;
            }
        }
        internal static void SetDrawingPropertiesFillBasic(this RenderItem item, ExcelTheme theme, ExcelDrawingFillBasic fill, ExcelDrawingColorManager color, bool gradientUserSpaceOnUse, Color? nullColor)
        {
            double? opacity = null;
            switch (fill.Style)
            {
                case eFillStyle.NoFill:
                    if (fill.IsEmpty) //Do NOT remove. This if is required for Shapes
                    {
                        item.FillColor = GetFillColor(theme, fill, color, item.FillColorSource, out opacity, nullColor);
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
                    item.GradientFill = new DrawingRenderGradientFill(theme, fill.GradientFill, gradientUserSpaceOnUse);
                    item.FillColor = null;
                    break;
            }
            if (opacity.HasValue)
            {
                item.FillOpacity = opacity;
            }
        }
        internal static void SetDrawingPropertiesBorder(this RenderItem item, ExcelTheme theme, ExcelDrawingBorder border, ExcelChartStyleColorManager color, bool hasBorder, Color? nullColor=null, double defaultWidth = 1.5, bool grandientUserSpaceOnUse=true)
        {
            double? opacity = null;
            if (border == null)
            {
                if (hasBorder)
                {
                    item.BorderColor = GetFillColor(theme, null, color, item.BorderColorSource, out opacity, nullColor ?? theme.ColorScheme.Dark1.GetColor());
                }
            }
            else
            {
                switch (border.Fill.Style)
                {
                    case eFillStyle.NoFill:
                        if (border.Fill.IsEmpty)
                        { 
                            item.BorderColor = GetFillColor(theme, border.Fill, color, item.BorderColorSource, out opacity, nullColor ?? theme.ColorScheme.Dark1.GetColor());
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
                        item.BorderGradientFill = new DrawingRenderGradientFill(theme, border.Fill.GradientFill, grandientUserSpaceOnUse);
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

        private static string GetFillColor(ExcelTheme theme, ExcelDrawingFillBasic fill, ExcelDrawingColorManager styleFillColor, PathFillMode fillColorSource, out double? opacity,  Color? nullColor = null)
        {
            opacity = null;
            if (fillColorSource == PathFillMode.None)
            {
                return "none";
            }

            Color fc;
            if (fill == null || fill.Style == eFillStyle.NoFill)
            {
                //Set fallback nullcolor
                if (nullColor != null && fill != null && fill.IsEmpty)
                {
                    fc = nullColor.Value;
                }


                if (styleFillColor == null)
                {
                    //There is no Style-Specified color. Themed Fill should be applied if it exists
                    //Fallback to theme
                    if (theme.FormatScheme.BackgroundFillStyle != null)
                    {
                        //Usually, at least for chart objects if the theme fill is not NoFill it is Subtle
                        var subtleBg = theme.FormatScheme.BackgroundFillStyle[0];
                        if (subtleBg.IsEmpty == false)
                        {
                            if (subtleBg.Style == eFillStyle.SolidFill)
                            {
                                if (subtleBg.SolidFill.Color.ColorType == eDrawingColorType.Scheme)
                                {
                                    //The theme color is PhClr which is fallback color to style.
                                    //Style does not exist.
                                    //But The base theme schemecolor does.
                                    //Hardcoded defaults to solid fill according to docs is Bg1
                                    //However since PhClr could also be a reference to StyleFillColor which does not exist.
                                    //Fallback to nullcolor as extra backup

                                    if (nullColor == null)
                                    {
                                        //return string.Empty;

                                        var bg1 = theme.ColorScheme.GetColorByEnum(eSchemeColor.Background1);
                                        fc = tc.ColorConverter.GetThemeColor(bg1);
                                    }
                                    else
                                    {
                                        fc = nullColor.Value;
                                    }
                                }
                                else 
                                {
                                    fc = subtleBg.Color;
                                }
                            }
                            else
                            {
                                //alternatively accent 1
                                fc = subtleBg.Color;
                            }
                        }
                        else
                        {
                            fc = subtleBg.Color;
                        }
                    }
                    else
                    {
                       return string.Empty;
                    }
                }
                else
                {
                    if (styleFillColor.ColorType == eDrawingColorType.Scheme)
                    {
                        var bg1 = theme.ColorScheme.GetColorByEnum(styleFillColor.SchemeColor.Color);
                        fc = bg1.GetColor();
                    }
                    else
                    {
                        fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill.Color, styleFillColor);
                    }
                }
            }
            else if (fill.Style == eFillStyle.SolidFill)
            {
                //Send in styleFill as well since a solid fill can refer to style color
                fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill.Color, styleFillColor);
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
