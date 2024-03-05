/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Packaging.Ionic;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Globalization;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using System.Text;
using System;

namespace OfficeOpenXml.Style
{
    public class ExcelRichTextColorNew
    {
        private ExcelRichTextNew _rt;

        internal ExcelRichTextColorNew(ExcelRichTextNew rt)
        {
            _rt = rt;
        }

        public ExcelRichTextColorNew(XmlReader xr)
        {
            int num;
            Auto = ConvertUtil.GetValueBool(xr.GetAttribute("auto"))??false;
            if (int.TryParse(xr.GetAttribute("indexed"), NumberStyles.Integer, CultureInfo.InvariantCulture, out num))
            {
                Indexed = num;
            }
            var rgb = xr.GetAttribute("rgb");
            if (!String.IsNullOrEmpty(rgb))
            {
                Rgb = ExcelDrawingRgbColor.GetColorFromString(rgb);
            }
            if (int.TryParse(xr.GetAttribute("theme"), NumberStyles.Integer, CultureInfo.InvariantCulture, out num))
            {
                Theme = (eThemeSchemeColor)num;
            }
            if(ConvertUtil.TryParseNumericString(xr.GetAttribute("tint"), out double d))
            {
                Tint = d;
            }
        }

        /// <summary>
        /// Gets the rgb color depending in <see cref="Rgb"/>, <see cref="Theme"/> and <see cref="Tint"/>
        /// </summary>
        public Color Color
        {
            get
            {
                Color ret = Color.Empty;
                if(Rgb != Color.Empty)
                {
                    ret = Rgb;
                }
                else if (Indexed.HasValue)
                {
                    ret = ExcelColor.GetIndexedColor(Indexed.Value);
                }
                else if(Theme.HasValue)
                {
                    ret = Utils.ColorConverter.GetThemeColor(_rt._collection._wb.ThemeManager.GetOrCreateTheme(), Theme.Value);
                }
                else if(Auto)
                {
                    ret = Color.Black;
                }
                if (Tint.HasValue)
                {
                    return Utils.ColorConverter.ApplyTint(ret, Tint.Value);
                }
                return ret;
            }
        }
        /// <summary>
        /// The rgb color value set in the file.
        /// </summary>
        public Color Rgb { get; set; }

        /// <summary>
        /// The color theme.
        /// </summary>
        public eThemeSchemeColor? Theme { get; set; }
        /// <summary>
        /// The tint value for the color.
        /// </summary>
        public double? Tint { get; set; }

        /// <summary>
        /// Auto color
        /// </summary>
        public bool Auto { get; set; }

        /// <summary>
        /// The indexed color number.
        /// A negative value means not set.
        /// </summary>
        public int? Indexed { get; set; }

        internal void AppendXml(StringBuilder sb)
        {
            sb.Append("<color");
            if(Auto )
            {
                sb.Append(" auto=\"1\"");
            }
            if( Indexed != null )
            {
                sb.Append($" indexed=\"{Indexed.Value}\"");
            }
            if (Rgb != Color.Empty)
            {
                sb.Append($" rgb=\"{(Rgb.ToArgb() & 0xFFFFFF).ToString("X").PadLeft(6, '0')}\"");
            }
            if(Theme != null)
            {
                sb.Append($" theme=\"{(int)Theme.Value}\"");
            }
            if(Tint != null)
            {
                sb.Append($" tint=\"{Tint.Value.ToString(CultureInfo.InvariantCulture)}\"");
            }
            sb.Append("/>");
        }
    }
}