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
using EPPlus.Export.Pdf.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Renderer.RenderItems.Fill;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Security.Cryptography.Xml;
using tc = OfficeOpenXml.Utils.TypeConversion;
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
        internal static void SetDrawingPropertiesFill(this RenderItem item, ExcelTheme theme, ExcelDrawingFill fill, ExcelDrawingColorManager color, UserSpaceSettings gradientUserSpace = UserSpaceSettings.ObjectBoundingBox, Color? nullColor = null)
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
        internal static void SetDrawingPropertiesFillBasic(this RenderItem item, ExcelTheme theme, ExcelDrawingFillBasic fill, ExcelDrawingColorManager color, UserSpaceSettings gradientUserSpaceOnUse, Color? nullColor)
        {
            double opacity = double.NaN;

            var fillNew = GetFillNew(fill, theme, color, item.FillColorSource, out opacity, () => { return nullColor; }, out DrawingRenderGradientFill gradFill);

            if(gradFill != null)
            {
                //Special case for gradFIll as it does not return string
                item.GradientFill = gradFill;
                item.FillType = FillType.GradientFill;
                item.FillColor = null;
            }
            else
            {
                item.FillColor = fillNew;
            }

            if (opacity != double.NaN)
            {
                item.FillOpacity = opacity;
            }

            //switch (fill.Style)
            //{
            //    case eFillStyle.NoFill:
            //        item.FillColor = GetFillNew(fill)
            //        //if (fill.IsEmpty) //Do NOT remove. This if is required for Shapes
            //        //{
            //        //    item.FillColor = GetFillColor(theme, fill, color, item.FillColorSource, out opacity, nullColor);
            //        //}
            //        //else
            //        //{
            //        //    item.FillColor = "none";
            //        //}
            //        break;
            //    case eFillStyle.SolidFill:
            //        item.FillColor = GetFillColor(theme, fill, color, item.FillColorSource, out opacity);
            //        break;
            //    case eFillStyle.GradientFill:
            //        item.GradientFill = new DrawingRenderGradientFill(theme, fill.GradientFill, gradientUserSpaceOnUse);
            //        item.FillType = FillType.GradientFill;
            //        item.FillColor = null;
            //        break;
            //}
            //if (opacity.HasValue)
            //{
            //    item.FillOpacity = opacity;
            //}
        }

        //bg1 is the hard-coded default of solid fill according to ooxml docs (MS-OE376)
        private static Color GetSchemeColor(ExcelTheme theme, eSchemeColor schemeColor = eSchemeColor.Background1)
        {
            var bg1 = theme.ColorScheme.GetColorByEnum(schemeColor);
            return tc.ColorConverter.GetThemeColor(bg1);
        }

        private static Color? GetFillColorFromTheme(ExcelTheme theme, Func<Color?> GetDefaultThemeColor)
        {
            Color? fc = GetDefaultThemeColor();

            if (fc.HasValue == false)
            {
                //Bg1 or alternatively accent 1
                fc = theme.FormatScheme.BackgroundFillStyle[0].Color;
            }
            return fc;
        }

        private static Color? GetFillColorFromReference(ExcelDrawingColorManager styleFillColor, ExcelTheme theme, ExcelDrawingFillBasic fill)
        {
            if(styleFillColor != null)
            {
                Color? fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill?.Color, styleFillColor);

                //if (styleFillColor.ColorType == eDrawingColorType.Scheme)
                //{
                //    var bg1 = theme.ColorScheme.GetColorByEnum(styleFillColor.SchemeColor.Color);
                //    fc = bg1.GetColor();
                //    var differentResultMB = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill?.Color, styleFillColor);
                //    //fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill?.Color, styleFillColor);
                //}
                //else
                //{
                //    if (fill != null && fill.Style != eFillStyle.NoFill)
                //    {
                //        fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill?.Color, styleFillColor);
                //    }
                //    else
                //    {
                //        return Color.Empty;
                //    }
                //}
            }
            return null;
        }

        private static string GetFallbackFill(ExcelTheme theme, ExcelDrawingFillBasic itemFill, ExcelDrawingColorManager reference, PathFillMode colorSource, out double opacity, Func<Color?> GetDefaultThemeColor)
        {
            Color? fc = null;

            //We already know the fill has "NoFill"
            //NoFill has two cases. Either the node does not exist. Or it has been set to NoFill specifically
            if (itemFill.IsEmpty)
            {
                //The node itself does not exist. It needs to check for potential fallbacks
                //Move on to 2. StyleManager
                fc = GetFillColorFromReference(reference, theme, itemFill);

                if (fc.HasValue == false)
                {
                    
                    //Move on to 3. Theme
                    fc = GetFillColorFromTheme(theme, GetDefaultThemeColor);

                }
            }
            else
            {
                opacity = 0d;
                //The node has specifically been set to NoFill AKA Transparent
                return "none";
            }

            if (fc.HasValue == false)
            {
                throw new InvalidOperationException("Fallback color must exist");
            }

            return GetAdjustmentsAndTransparency(fc.Value, colorSource, out opacity);
        }


        private static string GetAdjustmentsAndTransparency(Color fc, PathFillMode colorSource, out double opacity)
        {
            fc = tc.ColorConverter.GetAdjustedColor(colorSource, fc);
            if (fc.A < 255 && fc != Color.Empty)
            {
                opacity = fc.A / 255D;
            }
            else
            {
                opacity = 1d;
            }
            return "#" + fc.ToArgb().ToString("x8").Substring(2);
        }

        internal static string GetFillNew(ExcelDrawingFillBasic fill, ExcelTheme theme, ExcelDrawingColorManager reference, PathFillMode fillMode, out double opacity, Func<Color?> GetHardCodedDefaultForItem, out DrawingRenderGradientFill gradFill)
        {
            string fillStr = string.Empty;
            gradFill = null;
            opacity = 1d;

            //The Fallback chain of styles for drawing objects is:
            //1. Chart.Border (make sure to note the chart style ID
            //2. Chart.StyleManager.ChartArea.BorderReference
            //3. Theme.FormatScheme.BorderStyle[0] for subtle, [1] Moderate [2] Intense
            //4. If none of these contain even an empty node for the relevant property, Fallback to hardcoded documentation defaults 

            switch (fill.Style)
            {
                case eFillStyle.NoFill:
                    //Either transparent or Fallback to style hierarhy (options 2, 3 or 4)
                    fillStr = GetFallbackFill(theme, fill, reference, fillMode, out opacity, GetHardCodedDefaultForItem);
                    break;
                case eFillStyle.SolidFill:
                    //1. Standard case. There is a fill color to apply.
                    //Send in styleFill as well since a solid fill can refer to style color
                    var fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill.Color, reference);
                    fillStr = GetAdjustmentsAndTransparency(fc, fillMode, out opacity);
                    break;
                case eFillStyle.GradientFill:
                    gradFill = new DrawingRenderGradientFill(theme, fill.GradientFill, UserSpaceSettings.UserSpaceOnUse_Global);
                    break;
            }

            return fillStr;
        }

        internal static void SetDrawingBorderPropertiesNew(this RenderItem item, ExcelTheme theme, ExcelChartStyleColorManager reference, ExcelDrawingBorder border, double opacity, Func<Color?> GetHardCodedDefaultForItem)
        {
            var fillColorStr = GetFillNew(border.Fill, theme, reference, item.BorderColorSource, out opacity, GetHardCodedDefaultForItem, out DrawingRenderGradientFill gradFill);

            if(gradFill != null)
            {
                //Special case as gradfill does not return a string
                item.BorderGradientFill = new DrawingRenderGradientFill(theme, border.Fill.GradientFill, UserSpaceSettings.UserSpaceOnUse_Global);
                item.BorderColor = null;
            }
            else
            {
                item.BorderColor = fillColorStr;
                item.BorderGradientFill = null;
            }

            item.BorderOpacity = opacity;

            if (item.BorderColorSource != PathFillMode.None)
            {
                item.BorderWidth = (border?.Width ?? 0D) == 0D ? 0.75d : border.Width;
                if (border != null && border.LineStyle.HasValue && border.LineStyle != eLineStyle.Solid)
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

        internal static void SetDrawingPropertiesBorder(this RenderItem item, ExcelTheme theme, ExcelDrawingBorder border, ExcelChartStyleColorManager color, bool hasBorder, Color? nullColor=null, double defaultWidth = 1.5, UserSpaceSettings gradientUserSpaceOnUse = UserSpaceSettings.UserSpaceOnUse_Global, eChartStyle styleId = eChartStyle.Style2)
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
                        item.BorderGradientFill = new DrawingRenderGradientFill(theme, border.Fill.GradientFill, gradientUserSpaceOnUse);
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
                if(gc.A>0)
                {
                    item.GlowOpacity = Math.Round(gc.A / 255D * 100); 
                }
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

        private static string GetFillColor(ExcelTheme theme, ExcelDrawingFillBasic fill, ExcelDrawingColorManager styleFillColor, PathFillMode fillColorSource, out double? opacity,  Color? nullColor = null, eChartStyle chartStyle = eChartStyle.Style2)
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
                    //There is no Style-Specified color. Or rather. There is no styleSheet inside of the Chart folder. Themed Fill should be applied if it exists
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
                        if(fill != null && fill.Style != eFillStyle.NoFill)
                        {
                            fc = tc.ColorConverter.GetThemeColor(theme, fill.SolidFill?.Color, styleFillColor);
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                }
            }
            else if (fill.Style == eFillStyle.SolidFill)
            {
                fc = fill.Color;
                tc.ColorConverter.GetThemeColor(theme, fill.SolidFill.Color);
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
