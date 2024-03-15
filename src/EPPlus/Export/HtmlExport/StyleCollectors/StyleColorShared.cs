/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using System;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal static class StyleColorShared
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
