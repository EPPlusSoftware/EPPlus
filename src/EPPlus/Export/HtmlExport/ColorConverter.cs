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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using System.Drawing;
using OfficeOpenXml.Drawing.Theme;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class ColorConverter
    {
        internal static Color GetThemeColor(ExcelTheme theme, eThemeSchemeColor tc)
        {
            var cm = theme.ColorScheme.GetColorByEnum(tc);
            return GetThemeColor(cm);
        }
        private static Color GetThemeColor(ExcelDrawingThemeColorManager cm)
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
        }
    }
}
