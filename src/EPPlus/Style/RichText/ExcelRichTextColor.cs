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
    /// <summary>
    /// 
    /// </summary>
    public class ExcelRichTextColor
    {
        /// <summary>
        /// 
        /// </summary>
        public bool HasAttributes
        {
            get
            {
                return     Rgb != Color.Empty ||
                         Theme != null ||
                          Tint != null ||
                          Auto != null ||
                       Indexed != null;
            }
        }
        /// <summary>
        /// The rgb color value set in the file.
        /// </summary>
        public Color Rgb { get; set; } = Color.Empty;

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
        public bool? Auto { get; set; }

        /// <summary>
        /// The indexed color number.
        /// A negative value means not set.
        /// </summary>
        public int? Indexed { get; set; }

        internal ExcelRichTextColor Clone()
        {
            return new ExcelRichTextColor
            {
                Rgb = this.Rgb,
                Theme = this.Theme,
                Tint = this.Tint,
                Auto = this.Auto,
                Indexed = this.Indexed
            };
        }
    }
}