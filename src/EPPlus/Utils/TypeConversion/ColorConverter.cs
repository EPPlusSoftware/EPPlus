/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/15/2021         EPPlus Software AB       Html export
 *************************************************************************************************/
using EPPlus.DrawingRenderer;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Drawing;
using System.Linq;
using TC = OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Utils.TypeConversion
{
    public class ColorConverter
    {
        public static Color GetThemeColor(ExcelTheme theme, eThemeSchemeColor tc)
        {
            var cm = theme.ColorScheme.GetColorByEnum(tc);
            return GetThemeColor(cm);
        }
        public static Color GetThemeColor(ExcelTheme theme, ExcelDrawingColorManager cm)
        {
            if(cm!=null && cm.ColorType==eDrawingColorType.Scheme)
            {
                var newCm=theme.ColorScheme.GetColorByEnum(cm.SchemeColor.Color);
                if (newCm == null) return Color.Empty;
                var nc = GetThemeColor(newCm);
                return ApplyTransforms(nc, cm.Transforms);
            }
            var c=GetThemeColor(cm);
            return ApplyTransforms(c, cm.Transforms);

        }

        public static Color GetThemeColor(ExcelTheme theme, ExcelDrawingColorManager cm, ExcelDrawingColorManager cmStyle)
        {
            if (cm != null && cm.ColorType == eDrawingColorType.Scheme)
            {
                ExcelDrawingThemeColorManager newCm;
                if(cm.SchemeColor.Color == eSchemeColor.Style)
                {
                   return GetThemeColor(theme, cmStyle);
                }
                else
                {
                    newCm = theme.ColorScheme.GetColorByEnum(cm.SchemeColor.Color);
                }
                var nc = GetThemeColor(newCm);
                return ApplyTransforms(nc, cm.Transforms);
            }
            var c = GetThemeColor(cm);
            return ApplyTransforms(c, cm.Transforms);

        }

        internal static Color ApplyTransforms(Color c, ExcelColorTransformCollection transforms)
        {
            if (transforms==null || transforms.Count == 0) return c;

            var r = c.R;
            var g = c.G;
            var b = c.B;

            foreach (var t in transforms)
            {
                var v = t.Value / 100;
                switch(t.Type)
                {
                    case eColorTransformType.Shade:
                        c = ApplyTintDrawing(c, -(1-v));
                        break;
                    case eColorTransformType.Tint:
                        c = ApplyTintDrawing(c, v);
                        break;
                    case eColorTransformType.HueMod:
                        c = ApplyHueMod(c, v);
                        break;
                    case eColorTransformType.HueOff:
                        c = ApplyHueMod(c, 1, v);
                        break;
                    case eColorTransformType.SatMod:
                        c = ApplySatMod(c, v);
                        break;
                    case eColorTransformType.SatOff:
                        c = ApplySatMod(c, 1, v);
                        break;
                    case eColorTransformType.LumMod:
                        c = ApplyLumMod(c, v);
                        break;
                    case eColorTransformType.LumOff:
                        c = ApplyLumMod(c, 1, v);
                        break;
                    case eColorTransformType.Alpha:
                        c = Color.FromArgb((byte)Math.Round(255 * v), c.R, c.G, c.B);
                        break;
                    case eColorTransformType.AlphaMod:
                        c = Color.FromArgb((byte)Math.Round(c.A * v), c.R, c.G, c.B);
                        break;
                    case eColorTransformType.AlphaOff:
                        c = Color.FromArgb((byte)(c.A + v), c.R, c.G, c.B);
                        break;
                }
            }
            return c;
            //return Color.FromArgb(r, g, b);
        }
        internal static Color ApplyHueMod(Color c, double hueMod = 1, double hueOff = 0)
        {
            ExcelDrawingRgbColor.GetHslColor(c, out double h, out double s, out double l);

            h = Math.Max(0, Math.Min(1, l * hueMod + hueOff));
            var ret = ExcelDrawingHslColor.GetRgb(h, s, l);
            return ret;
        }

        internal static Color ApplyLumMod(Color c, double lumMod=1, double lumOff=0)
        {
            ExcelDrawingRgbColor.GetHslColor(c, out double h, out double s, out double l);

            l = Math.Max(0, Math.Min(1, l * lumMod + lumOff));
            var ret = ExcelDrawingHslColor.GetRgb(h, s, l);
            return ret;
        }
        internal static Color ApplySatMod(Color c, double satMod = 1, double satOff = 0)
        {
            var h = c.GetHue();
            var s = c.GetSaturation();
            var l = c.GetBrightness();

            ExcelDrawingRgbColor.GetHslColor(c, out double h2, out double s2, out double l2);

            var ret1 = ExcelDrawingHslColor.GetRgb(h, s* satMod + satOff, l);
            var ret2 = ExcelDrawingHslColor.GetRgb(h2, s2 * satMod + satOff, l2);
            return ret2;
        }

        /// <summary>
        /// Converts the color to a <see cref="Color"/>
        /// </summary>
        /// <param name="cm">The theme color manager</param>
        /// <returns>The RGB color</returns>
        public static Color GetThemeColor(ExcelDrawingThemeColorManager cm)
        {
            Color color;
            switch (cm.ColorType)
            {
                case eDrawingColorType.Rgb:
                    color = cm.RgbColor.Color;
                    break;
                case eDrawingColorType.Preset:
                    color = Color.FromName(cm.PresetColor.Color.ToString());
                    break;
                case eDrawingColorType.System:
                    color = cm.SystemColor.GetColor();
                    break;
                case eDrawingColorType.RgbPercentage:
                    var rp = cm.RgbPercentageColor;
                    color = Color.FromArgb(GetRgpPercentToRgb(rp.RedPercentage),
                                           GetRgpPercentToRgb(rp.GreenPercentage),
                                           GetRgpPercentToRgb(rp.BluePercentage));
                    
                    break;
                case eDrawingColorType.Hsl:
                    color = cm.HslColor.GetRgbColor();
                    break;
                default:
                    color = Color.Empty;
                    break;
            }

            //TODO:Apply Transforms
            return color;
        }

        private static int GetRgpPercentToRgb(double percentage)
        {
            if (percentage < 0) return 0;
            if (percentage > 255) return 255;
            return (int)(percentage * 255 / 100);
        }
        internal static Color ApplyTint(Color ret, double tint)
        {
            if (tint == 0)
            {
                return ret;
            }
            else
            {
                ExcelDrawingRgbColor.GetHslColor(ret, out double h, out double s, out double l);
                if (tint < 0)
                {
                    l = l * (1.0 + tint);
                }
                else if (tint > 0)
                {
                    l += (1 - l) * tint;
                }
                return ExcelDrawingHslColor.GetRgb(h, s, l);
            }
            //if (tint < 0)
            //{
            //    double shade = 1+tint;
            //    var r = (byte)Math.Round(ret.R * shade);
            //    var g = (byte)Math.Round(ret.G * shade);
            //    var b = (byte)Math.Round(ret.B * shade); 
            //    return Color.FromArgb(ret.A, r, g, b);
            //}
            //else if(tint > 0)
            //{
            //    double blend = 1.0 - tint;
            //    var r = (byte)Math.Round(ret.R + (255 - ret.R) * blend);
            //    var g = (byte)Math.Round(ret.G + (255 - ret.G) * blend);
            //    var b = (byte)Math.Round(ret.B + (255 - ret.B) * blend);
            //    return Color.FromArgb(ret.A, r, g, b);
            //}
            //return ret;
        }
        internal static Color ApplyTintDrawing(Color ret, double tint)
        {
            //if (tint == 0)
            //{
            //    return ret;
            //}
            //else
            //{
            //    ExcelDrawingRgbColor.GetHslColor(ret, out double h, out double s, out double l);
            //    if (tint < 0)
            //    {
            //        l = l * (1.0 + tint);
            //    }
            //    else if (tint > 0)
            //    {
            //        l += (1 - l) * tint;
            //    }
            //    return ExcelDrawingHslColor.GetRgb(h, s, l);
            //}
            if (tint < 0)
            {
                double shade = 1 + tint;
                var r = (byte)Math.Round(ret.R * shade);
                var g = (byte)Math.Round(ret.G * shade);
                var b = (byte)Math.Round(ret.B * shade);
                return Color.FromArgb(ret.A, r, g, b);
            }
            else if (tint > 0)
            {
                double blend = 1.0 - tint;
                var r = (byte)Math.Round(ret.R + (255 - ret.R) * blend);
                var g = (byte)Math.Round(ret.G + (255 - ret.G) * blend);
                var b = (byte)Math.Round(ret.B + (255 - ret.B) * blend);
                return Color.FromArgb(ret.A, r, g, b);
            }
            return ret;
        }

        internal static Color ApplyBlend(Color color, Color blendColor, double percent)
        {
            var colorPercent = 1 - percent;
            var r = (int)Math.Min(255D, color.R * colorPercent + blendColor.R * percent);
            var g = (int)Math.Min(255D, color.G * colorPercent + blendColor.G * percent);
            var b = (int)Math.Min(255D, color.B * colorPercent + blendColor.B * percent);
            return Color.FromArgb(0xff, r, g, b);
        }
        internal static Color GetAdjustedColor(PathFillMode fillColorSource, Color fc)
        {
            switch (fillColorSource)
            {
                case PathFillMode.Darken:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.Black, 0.4);
                    break;
                case PathFillMode.DarkenLess:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.Black, 50D / 255D);
                    break;
                case PathFillMode.LightenLess:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.White, 50D / 255D);
                    break;
                case PathFillMode.Lighten:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.White, 0.4);
                    break;
            }

            return fc;
        }

        internal static double GetOpacity(ExcelDrawingColorManager color)
        {
            if(color.Transforms.Where(t=>t.Type==eColorTransformType.Alpha).FirstOrDefault() is IColorTransformItem alpha)
            {
                return alpha.Value / 100;
            }
            return 1D;
        }
    }
}
