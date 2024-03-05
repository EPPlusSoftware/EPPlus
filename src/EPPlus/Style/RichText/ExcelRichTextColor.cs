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

namespace OfficeOpenXml.Style
{
    public class ExcelRichTextColor : XmlHelper
    {
        private ExcelRichText _rt;

        internal ExcelRichTextColor(XmlNamespaceManager ns, XmlNode topNode, ExcelRichText rt) : base(ns, topNode)
        {
            _rt = rt;
        }
        /// <summary>
        /// Gets the rgb color depending in <see cref="Rgb"/>, <see cref="Theme"/> and <see cref="Tint"/>
        /// </summary>
        public Color Color
        {
            get
            {
                return _rt.Color;
            }
        }
        /// <summary>
        /// The rgb color value set in the file.
        /// </summary>
        public Color Rgb
        {
            get
            {
                var col = GetXmlNodeString(ExcelRichText.COLOR_PATH);
                if (string.IsNullOrEmpty(col))
                {
                    return Color.Empty;
                }
                return Color.FromArgb(int.Parse(col, NumberStyles.AllowHexSpecifier));
            }
            set
            {
                _rt._collection.ConvertRichtext();
                if (value==Color.Empty)
                {
                    DeleteNode(ExcelRichText.COLOR_PATH);
                }
                else
                {
                    SetXmlNodeString(ExcelRichText.COLOR_PATH, value.ToArgb().ToString("X"));
                }
                if (_rt._callback != null) _rt._callback();
            }
        }
        /// <summary>
        /// The color theme.
        /// </summary>
        public eThemeSchemeColor? Theme
        {
            get
            {
                return GetXmlNodeString(ExcelRichText.COLOR_THEME_PATH).ToEnum<eThemeSchemeColor>();
            }
            set
            {
                _rt._collection.ConvertRichtext();
                var v =value.ToEnumString();
                if(v==null)
                {
                    DeleteNode(ExcelRichText.COLOR_THEME_PATH);
                }
                else
                {
                    SetXmlNodeString(ExcelRichText.COLOR_THEME_PATH, v);
                }
                if (_rt._callback != null) _rt._callback();
            }
        }
        /// <summary>
        /// The tint value for the color.
        /// </summary>
        public double? Tint
        {
            get
            {
                return GetXmlNodeDoubleNull(ExcelRichText.COLOR_TINT_PATH);
            }
            set
            {
                _rt._collection.ConvertRichtext();
                SetXmlNodeDouble(ExcelRichText.COLOR_TINT_PATH, value, true);
                if (_rt._callback != null) _rt._callback();
            }
        }
    }
}