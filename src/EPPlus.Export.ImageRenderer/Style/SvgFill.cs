using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;

namespace EPPlus.Export.ImageRenderer.Style
{
    internal class SvgFill : IFillBasic
    {
        ExcelDrawingFillBasic _fill;

        public SvgFill(ExcelDrawingFillBasic fill)
        {
            Style = fill.Style;
            _fill = fill;
        }

        public SvgFill(ExcelDrawingFill fill)
        {
            Style = fill.Style;
            _fill = fill;
        }

        public eFillStyle Style {  get; set; }

        public bool HasValue
        {
            get
            {
                return !_fill.IsEmpty;
            }
        }


        public bool IsGradient
        {
            get
            {
                return Style == eFillStyle.GradientFill;
            }
        }

        public string GetBackgroundColor(ExcelTheme theme)
        {
            return GetColor(_fill.Color, theme);
        }

        /// <summary>
        /// Gets hexcode color for html as a string 
        /// </summary>
        /// <param name="c"></param>
        /// <param name="theme"></param>
        internal static string GetColor(Color c, ExcelTheme theme)
        {
            //Color ret;
            //if (!string.IsNullOrEmpty(c.ToColorString()))
            //{
            //    if (int.TryParse(c.Rgb, NumberStyles.HexNumber, null, out int hex))
            //    {
            //        ret = Color.FromArgb(hex);
            //    }
            //    else
            //    {
            //        ret = Color.Empty;
            //    }
            //}
            //else if (c.Theme.HasValue)
            //{
            //    ret = Utils.TypeConversion.ColorConverter.GetThemeColor(theme, c.Theme.Value);
            //}
            //else if (c.Indexed >= 0)
            //{
            //    ret = theme._wb.Styles.GetIndexedColor(c.Indexed);
            //}
            //else
            //{
            //    //Automatic, set to black.
            //    if (c.Auto)
            //    {
            //        ret = Color.Black;
            //    }
            //    else if (c.Exists)
            //    {
            //        ret = Color.Empty;
            //    }
            //    else
            //    {
            //        return null;
            //    }
            //}
            //if (c.Tint != 0)
            //{
            //    ret = Utils.TypeConversion.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            //}

            return "#" + c.ToArgb().ToString("x8").Substring(2);
        }

        public string GetGradientColor1(ExcelTheme theme)
        {
            return GetColor(_fill.GradientFill.Colors.ToArray()[0].Color.GetColor(), theme);
        }
        public string GetGradientColor2(ExcelTheme theme)
        {
            return GetColor(_fill.GradientFill.Colors.ToArray()[1].Color.GetColor(), theme);
        }

    }
}
