using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class StyleColorXml :IStyleColor
    {
        ExcelColor _color;

        public StyleColorXml(ExcelColor color) 
        {
            _color = color;
        }

        public bool Exists { get; }

        public bool Auto { get; }

        public int Indexed { get; }

        public double Tint { get; }

        public eThemeSchemeColor? Theme { get; }

        string Rgb { get; }

        string IStyleColor.Rgb { get; }

        public bool AreColorEqual(IStyleColor color)
        {
            if (Tint != color.Tint) return false;
            if (Indexed >= 0)
            {
                return Indexed == color.Indexed;
            }
            else if (string.IsNullOrEmpty(Rgb) == false)
            {
                return Rgb == color.Rgb;
            }
            else if (Theme != null)
            {
                return Theme == color.Theme;
            }
            else
            {
                return Auto == color.Auto;
            }
        }


        public string GetColor(ExcelTheme theme)
        {
            Color ret;
            if (!string.IsNullOrEmpty(Rgb))
            {
                if (int.TryParse(Rgb, NumberStyles.HexNumber, null, out int hex))
                {
                    ret = Color.FromArgb(hex);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else if (Theme.HasValue)
            {
                ret = Utils.ColorConverter.GetThemeColor(theme, Theme.Value);
            }
            else if (Indexed >= 0)
            {
                ret = ExcelColor.GetIndexedColor(Indexed);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (Tint != 0)
            {
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
    }
}
