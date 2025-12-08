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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Font;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// Defines a Theme within the package
    /// </summary>
    public class ExcelTheme : ExcelThemeBase
    {
        internal ExcelWorkbook _wb;
        /// <summary>
        /// The name of the theme
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }

        internal ExcelTheme(ExcelWorkbook workbook, ZipPackageRelationship rel)
            : base(workbook._package,workbook.NameSpaceManager, rel, "a:themeElements/")
        {
            _wb = workbook;
        }

        internal string GetFontByCode(string fontName)
        {
            string fontNameTheme;
            switch (fontName)
            {
                case "+mj-lt":
                    fontNameTheme = FontScheme.MajorFont.FirstOrDefault(x=>(x as ExcelDrawingFontSpecial)?.Type==eFontType.Latin)?.Typeface ?? "Aptos Display";
                    break;
                case "+mn-lt":
                    fontNameTheme = FontScheme.MinorFont.FirstOrDefault(x => (x as ExcelDrawingFontSpecial)?.Type == eFontType.Latin)?.Typeface ?? "Aptos Narrow";
                    break;
                case "+mj-ea":
                    fontNameTheme = FontScheme.MajorFont.FirstOrDefault(x => (x as ExcelDrawingFontSpecial)?.Type == eFontType.EastAsian)?.Typeface;
                    if (string.IsNullOrEmpty(fontNameTheme)) return GetFontByCode("+mj-lt");
                    break;
                case "+mn-ea":
                    fontNameTheme = FontScheme.MinorFont.FirstOrDefault(x => (x as ExcelDrawingFontSpecial)?.Type == eFontType.EastAsian)?.Typeface;
                    if (string.IsNullOrEmpty(fontNameTheme)) return GetFontByCode("+mn-lt");
                    break;
                case "+mj-cs":
                    fontNameTheme = FontScheme.MajorFont.FirstOrDefault(x => (x as ExcelDrawingFontSpecial)?.Type == eFontType.Complex)?.Typeface;
                    if (string.IsNullOrEmpty(fontNameTheme)) return GetFontByCode("+mj-lt");
                    break;
                case "+mn-cs":
                    fontNameTheme = FontScheme.MinorFont.FirstOrDefault(x => (x as ExcelDrawingFontSpecial)?.Type == eFontType.Complex)?.Typeface;
                    if (string.IsNullOrEmpty(fontNameTheme)) return GetFontByCode("+mn-lt");
                    break;
                default:
                    return fontName;
            }
            return fontNameTheme;
        }
    }
}
