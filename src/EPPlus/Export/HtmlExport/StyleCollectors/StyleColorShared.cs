using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    static class StyleColorShared
    {
        public static bool AreColorEqual(IStyleColor c1, IStyleColor c2)
        {
            if (c1.Tint != c2.Tint) return false;
            if (c1.Indexed >= 0)
            {
                return c1.Indexed == c2.Indexed;
            }
            else if (string.IsNullOrEmpty(c1.Rgb) == false)
            {
                return c1.Rgb == c2.Rgb;
            }
            else if (c1.Theme != null)
            {
                return c1.Theme == c2.Theme;
            }
            else
            {
                return c1.Auto == c2.Auto;
            }
        }

        public static string GetColor(IStyleColor color, ExcelTheme theme)
        {
            Color ret;
            if (!string.IsNullOrEmpty(color.Rgb))
            {
                if (int.TryParse(color.Rgb, NumberStyles.HexNumber, null, out int hex))
                {
                    ret = Color.FromArgb(hex);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else if (color.Theme.HasValue)
            {
                ret = Utils.ColorConverter.GetThemeColor(theme, color.Theme.Value);
            }
            else if (color.Indexed >= 0)
            {
                ret = theme._wb.Styles.GetIndexedColor(color.Indexed);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (color.Tint != 0)
            {
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(color.Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
    }
}
