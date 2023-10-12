using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class GenericColor
    {
        string _rgb = null;
        int _indexed = -1;
        eThemeSchemeColor? _theme;
        decimal _tint;
        
        internal GenericColor(ExcelColorXml color)
        {
            _rgb = color.Rgb;
            _indexed = color.Indexed;
            _theme = color.Theme;
            _tint = color.Tint;
        }

        internal GenericColor(ExcelDxfColor color)
        {
            _rgb = color.HasValue ? null : color.Color.Value.ToArgb().ToString();
            _indexed = color.Index.Value;
            _theme = color.Theme;
            _tint = (decimal)color.Tint.Value;
        }

        /// <summary>
        /// Gets hexcode color for html as a string 
        /// </summary>
        /// <param name="c"></param>
        /// <param name="theme"></param>
        /// <returns></returns>
        internal string GetHexCodeColor(ExcelTheme theme)
        {
            Color ret;
            if (!string.IsNullOrEmpty(_rgb))
            {
                if (int.TryParse(_rgb, NumberStyles.HexNumber, null, out int hex))
                {
                    ret = Color.FromArgb(hex);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else if (_theme.HasValue)
            {
                ret = Utils.ColorConverter.GetThemeColor(theme, _theme.Value);
            }
            else if (_indexed >= 0)
            {
                ret = ExcelColor.GetIndexedColor(_indexed);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (_tint != 0)
            {
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
    }
}
