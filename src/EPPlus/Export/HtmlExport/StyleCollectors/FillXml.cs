using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FillXml : IFill
    {
        ExcelFillXml _fill;

        internal FillXml(ExcelFillXml fill)
        {
            _fill = fill;
        }

        public ExcelFillStyle PatternType 
        { 
            get { return _fill.PatternType; } 
        }

        public bool IsGradient
        {
            get
            {
                return _fill is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None;
            }
        }

        public double Degree
        {
            get 
            {
                if (IsGradient)
                {
                    return ((ExcelGradientFillXml)_fill).Degree;
                }

                return double.NaN;
            }
        }

        public double Right
        {
            get
            {
                if (IsGradient)
                {
                    return ((ExcelGradientFillXml)_fill).Right;
                }

                return double.NaN;
            }
        }

        public double Bottom
        {
            get
            {
                if (IsGradient)
                {
                    return ((ExcelGradientFillXml)_fill).Bottom;
                }

                return double.NaN;
            }
        }

        public bool IsLinear
        {
            get
            {
                if (IsGradient)
                {
                    return ((ExcelGradientFillXml)_fill).Type == ExcelFillGradientType.Linear;
                }

                return false;
            }
        }

        public bool HasValue
        {
            get
            {
                return !string.IsNullOrEmpty(_fill.Id);
            }
        }

        public string GetBackgroundColor(ExcelTheme theme)
        {
            return GetColor(_fill.BackgroundColor, theme);
        }

        public string GetPatternColor(ExcelTheme theme)
        {
            return GetColor(_fill.PatternColor, theme);
        }

        public string GetGradientColor1(ExcelTheme theme)
        {
            return GetColor(((ExcelGradientFillXml)_fill).GradientColor1, theme);
        }
        public string GetGradientColor2(ExcelTheme theme)
        {
            return GetColor(((ExcelGradientFillXml)_fill).GradientColor2, theme);
        }

        /// <summary>
        /// Gets hexcode color for html as a string 
        /// </summary>
        /// <param name="c"></param>
        /// <param name="theme"></param>
        /// <returns></returns>
        internal static string GetColor(ExcelColorXml c, ExcelTheme theme)
        {
            Color ret;
            if (!string.IsNullOrEmpty(c.Rgb))
            {
                if (int.TryParse(c.Rgb, NumberStyles.HexNumber, null, out int hex))
                {
                    ret = Color.FromArgb(hex);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else if (c.Theme.HasValue)
            {
                ret = Utils.ColorConverter.GetThemeColor(theme, c.Theme.Value);
            }
            else if (c.Indexed >= 0)
            {
                ret = theme._wb.Styles.GetIndexedColor(c.Indexed);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (c.Tint != 0)
            {
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
    }
}
