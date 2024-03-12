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
using System.Text;
using System.Xml;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using System.Xml.Linq;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// 
    /// </summary>
    internal class ExcelRichTextAttributes
    {
        /// <summary>
        /// 
        /// </summary>
        public ExcelRichTextAttributes()
        {
            ColorAttributes = new ExcelRichTextColor();
        }

        /// <summary>
        /// Preserves whitespace. Default true
        /// </summary>
        public bool PreserveSpace { get; set; } = true;

        /// <summary>
        /// Bold text
        /// </summary>
        public bool? Bold { get; set; } = false;

        /// <summary>
        /// Italic text
        /// </summary>
        public bool? Italic { get; set; } = false;

        /// <summary>
        /// Strike-out text
        /// </summary>
        public bool? Strike { get; set; } = false;

        /// <summary>
        /// Vertical Alignment
        /// </summary>
        public ExcelVerticalAlignmentFont? VerticalAlign { get; set; } = ExcelVerticalAlignmentFont.None;

        /// <summary>
        /// Font size
        /// </summary>
        public float Size { get; set; } = 0f;

        /// <summary>
        /// Name of the font
        /// </summary>
        public string FontName { get; set; } = string.Empty;

        /// <summary>
        /// Color settings.
        /// <seealso cref="Color"/>
        /// </summary>
        public ExcelRichTextColor ColorAttributes { get; set; }

        /// <summary>
        /// Characterset to use
        /// </summary>
        public int? Charset { get; set; }

        /// <summary>
        /// Font family
        /// </summary>
        public int? Family { get; set; }

        /// <summary>
        /// Underline type of text
        /// </summary>
        public ExcelUnderLineType? UnderLineType { get; set; }

        //NOT SUPPOERTED
        ///// <summary>
        ///// Scheme of the text
        ///// </summary>
        //public eThemeFontCollectionType? Scheme { get; set; }

        ///// <summary>
        ///// Outline the text
        ///// </summary>
        //public bool Outline { get; set; }

        ///// <summary>
        ///// Apply shadow to text
        ///// </summary>
        //public bool Shadow { get; set; }

        ///// <summary>
        ///// condense the text
        ///// </summary>
        //public bool Condense { get; set; }

        ///// <summary>
        ///// Extend the text
        ///// </summary>
        //public bool Extend { get; set; }
    }

    /// <summary>
    /// A richtext part
    /// </summary>
    public class ExcelRichText
    {
        /// <summary>
        /// A referens to the richtext collection
        /// </summary>
        public ExcelRichTextCollection _collection { get; set; }

        #region RichText Properties Attributes
        /// <summary>
        /// Preserves whitespace. Default true
        /// </summary>
        public bool PreserveSpace { get; set; } = true;

        /// <summary>
        /// Bold text
        /// </summary>
        public bool? Bold { get => Attributes.Bold; set => Attributes.Bold = value; }

        /// <summary>
        /// Italic text
        /// </summary>
        public bool? Italic { get => Attributes.Italic; set => Attributes.Italic = value; }

        /// <summary>
        /// Strike-out text
        /// </summary>
        public bool? Strike { get => Attributes.Strike; set => Attributes.Strike = value; }

        /// <summary>
        /// Underlined text
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return Attributes.UnderLineType != null && Attributes.UnderLineType != ExcelUnderLineType.None;
            }
            set
            {
                Attributes.UnderLineType = value ? ExcelUnderLineType.Single : null;
            }
        }

        /// <summary>
        /// Vertical Alignment
        /// </summary>
        public ExcelVerticalAlignmentFont? VerticalAlign { get => Attributes.VerticalAlign; set => Attributes.VerticalAlign = value; }

        /// <summary>
        /// Font size
        /// </summary>
        public float Size { get => Attributes.Size; set => Attributes.Size = value; }

        /// <summary>
        /// Name of the font
        /// </summary>
        public string FontName { get => Attributes.FontName; set => Attributes.FontName= value; }


        /// <summary>
        /// Text color.
        /// Also see <seealso cref="ColorSettings"/>
        /// </summary>
        public Color Color
        {
            get
            {
                Color ret = Color.Empty;
                if (Attributes.ColorAttributes.Rgb != Color.Empty)
                {
                    ret = Attributes.ColorAttributes.Rgb;
                }
                else if (Attributes.ColorAttributes.Indexed.HasValue)
                {
                    ret = ExcelColor.GetIndexedColor(Attributes.ColorAttributes.Indexed.Value);
                }
                else if (Attributes.ColorAttributes.Theme.HasValue)
                {
                    ret = Utils.ColorConverter.GetThemeColor(_collection._wb.ThemeManager.GetOrCreateTheme(), Attributes.ColorAttributes.Theme.Value);
                }
                else if (Attributes.ColorAttributes.Auto)
                {
                    ret = Color.Black;
                }
                if (Attributes.ColorAttributes.Tint.HasValue)
                {
                    return Utils.ColorConverter.ApplyTint(ret, Attributes.ColorAttributes.Tint.Value);
                }
                return ret;
            }
            set
            {
                Attributes.ColorAttributes.Rgb = value;
            }
        }

        /// <summary>
        /// Color settings.
        /// <seealso cref="Color"/>
        /// </summary>
        public ExcelRichTextColor ColorSettings { get => Attributes.ColorAttributes; set => Attributes.ColorAttributes = value; }

        /// <summary>
        /// Characterset to use
        /// </summary>
        public int? Charset { get => Attributes.Charset; set => Attributes.Charset = value; }

        /// <summary>
        /// Font family
        /// </summary>
        public int? Family { get => Attributes.Family; set => Attributes.Family = value; }

        /// <summary>
        /// Underline type of text
        /// </summary>
        public ExcelUnderLineType? UnderLineType { get => Attributes.UnderLineType; set => Attributes.UnderLineType = value; }

        //NOT SUPPOERTED
        ///// <summary>
        ///// Scheme of the text
        ///// </summary>
        //public eThemeFontCollectionType? Scheme { get; set; }

        ///// <summary>
        ///// Outline the text
        ///// </summary>
        //public bool Outline { get; set; }

        ///// <summary>
        ///// Apply shadow to text
        ///// </summary>
        //public bool Shadow { get; set; }

        ///// <summary>
        ///// condense the text
        ///// </summary>
        //public bool Condense { get; set; }

        ///// <summary>
        ///// Extend the text
        ///// </summary>
        //public bool Extend { get; set; }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        internal ExcelRichTextAttributes Attributes;
        private string _text;

        /// <summary>
        /// The text
        /// </summary>
        public string Text
        {
            get
            {
                return _text;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new InvalidOperationException("Text can't be null or empty");
                }
                _text = value;
            }
        }

        internal ExcelRichText(string text, ExcelRichTextCollection collection)
        {
            Text = text;
            Attributes = new ExcelRichTextAttributes();
            _collection = collection;
            //ColorSettings = new ExcelRichTextColor();
        }

        internal ExcelRichText(string text, ExcelRichTextAttributes attributes, ExcelRichTextCollection collection)
        {
            Text = text;
            _collection = collection;
            Attributes = attributes;
        }

        /// <summary>
        /// Get the underline typ for rich text
        /// </summary>
        /// <param name="v"></param>
        /// <returns>returns excelunderline type</returns>
        public static ExcelUnderLineType GetUnderlineType(string v)
        {
            switch (v)
            {
                case "single":
                    return ExcelUnderLineType.Single;
                case "double":
                    return ExcelUnderLineType.Double;
                case "singleAccounting":
                    return ExcelUnderLineType.SingleAccounting;
                case "doubleAccounting":
                    return ExcelUnderLineType.DoubleAccounting;
                default:
                    return ExcelUnderLineType.Single;
            }
        }

        /// <summary>
        /// Get the underline typ for rich text
        /// </summary>
        /// <param name="v"></param>
        /// <returns>returns excelunderline type</returns>
        public static ExcelVerticalAlignmentFont GetUVerticalAlignmentFont(string v)
        {
            switch (v)
            {
                case "baseline":
                    return ExcelVerticalAlignmentFont.Baseline;
                case "subscript":
                    return ExcelVerticalAlignmentFont.Subscript;
                case "superscript":
                    return ExcelVerticalAlignmentFont.Superscript;
                default:
                    return ExcelVerticalAlignmentFont.None;
            }
        }

        string ValueHasWhiteSpaces()
        {
            if (Text != null && Text.Length > 0)
            {
                if (char.IsWhiteSpace(Text[0]) || char.IsWhiteSpace(Text[Text.Length - 1]))
                {
                    return " xml:space=\"preserve\"";
                }
            }
            return "";
        }

        /// <summary>
        /// Returns the rich text item as a html string.
        /// </summary>
        public string HtmlText
        {
            get
            {
                var sb = new StringBuilder();
                WriteHtmlText(sb);
                return sb.ToString();
            }
        }

        internal void WriteHtmlText(StringBuilder sb)
        {
            sb.Append("<span style=\"");
            HtmlRichText.GetRichTextStyle(this, sb);
            sb.Append("\">");
            sb.Append(Text);
            sb.Append("</span>");
        }

        /// <summary>
        /// Read RichText attributes from xml.
        /// </summary>
        /// <param name="xr"></param>
        internal static ExcelRichTextAttributes ReadrPr(XmlReader xr)
        {
            ExcelRichTextAttributes attributes = new();
            while (xr.Read())
            {
                if (xr.LocalName == "rPr") break;

                switch (xr.LocalName)
                {
                    case "b":
                        attributes.Bold = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "i":
                        attributes.Italic = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "strike":
                        attributes.Strike = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "u":
                        attributes.UnderLineType = GetUnderlineType(xr.GetAttribute("val"));
                        break;
                    case "vertAlign":
                        attributes.VerticalAlign = xr.GetAttribute("val").ToEnum<ExcelVerticalAlignmentFont>(ExcelVerticalAlignmentFont.None);
                        break;
                    case "sz":
                        if (ConvertUtil.TryParseNumericString(xr.GetAttribute("val"), out double num))
                        {
                            attributes.Size = Convert.ToSingle(num);
                        }
                        break;
                    case "rFont":
                        attributes.FontName = xr.GetAttribute("val");
                        break;
                    case "charset":
                        attributes.Charset = int.Parse(xr.GetAttribute("val"));
                        break;
                    case "family":
                        attributes.Family = int.Parse(xr.GetAttribute("val"));
                        break;
                    case "color":
                        attributes.ColorAttributes = ReadColor(xr);
                        break;
                        //case "outline":
                        //    Outline = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        //    break;
                        //case "shadow":
                        //    Shadow = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        //    break;
                        //case "condense":
                        //    Condense = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        //    break;
                        //case "extend":
                        //    Extend = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        //    break;
                        //case "scheme":
                        //    Scheme = xr.GetAttribute("val").ToEnum<eThemeFontCollectionType>(eThemeFontCollectionType.None);
                        //    break;
                }
            }
            return attributes;
        }

        private static ExcelRichTextColor ReadColor(XmlReader xr)
        {
            ExcelRichTextColor colorAttributes = new ExcelRichTextColor();
            int num;
            var auto = xr.GetAttribute("auto");
            if (int.TryParse(auto, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
            {
                colorAttributes.Auto = result > 0 || result < 0 ? true : false;
            }
            if (int.TryParse(xr.GetAttribute("indexed"), NumberStyles.Integer, CultureInfo.InvariantCulture, out num))
            {
                colorAttributes.Indexed = num;
            }
            var rgb = xr.GetAttribute("rgb");
            if (!String.IsNullOrEmpty(rgb))
            {
                colorAttributes.Rgb = ExcelDrawingRgbColor.GetColorFromString(rgb);
            }
            if (int.TryParse(xr.GetAttribute("theme"), NumberStyles.Integer, CultureInfo.InvariantCulture, out num))
            {
                colorAttributes.Theme = (eThemeSchemeColor)num;
            }
            var tint = xr.GetAttribute("tint");
            if (ConvertUtil.TryParseNumericString(tint, out double d, CultureInfo.InvariantCulture))
            {
                colorAttributes.Tint = d;
            }
            return colorAttributes;
        }

        /// <summary>
        /// Write RichTextAttributes
        /// </summary>
        /// <param name="sb"></param>
        internal void WriteRichTextAttributes(StringBuilder sb)
        {
            sb.Append("<r>");
            if (!HasDefaultValue)
            {
                sb.Append("<rPr>");
                if (!String.IsNullOrEmpty(FontName))
                {
                    sb.Append($"<rFont val=\"{FontName}\"/>");
                }
                if (Charset.HasValue)
                {
                    sb.Append($"<charset val=\"{Charset}\"/>");
                }
                if (Family.HasValue)
                {
                    sb.Append($"<family val=\"{Family}\"/>");
                }
                if (Bold != null && Bold == true)
                {
                    sb.Append($"<b/>");
                }
                if (Italic != null && Italic == true)
                {
                    sb.Append($"<i/>");
                }
                if (Strike != null && Strike == true)
                {
                    sb.Append($"<strike/>");
                }
                if (Color != Color.Empty)
                {
                    WriteRichTextColorAttributes(sb);
                }
                if (Size > 0)
                {
                    sb.Append($"<sz val=\"{Size.ToString(CultureInfo.InvariantCulture)}\"/>");
                }
                if (UnderLine)
                {
                    sb.Append($"<u val=\"{UnderLineType.Value.ToEnumString()}\"/>");
                }
                if (VerticalAlign != ExcelVerticalAlignmentFont.None)
                {
                    sb.Append($"<vertAlign val=\"{VerticalAlign.ToEnumString()}\"/>");
                }
                //NOT SUPPORTED
                //if (Outline)
                //{
                //    sb.Append($"<outline/>");
                //}
                //if (Shadow)
                //{
                //    sb.Append($"<shadow/>");
                //}
                //if (Condense)
                //{
                //    sb.Append($"<condense/>");
                //}
                //if (Extend)
                //{
                //    sb.Append($"<extend/>");
                //}
                //if (Scheme != null && Scheme != eThemeFontCollectionType.None)
                //{
                //    sb.Append($"<scheme val=\"{Scheme.ToEnumString()}\"/>");
                //}
                sb.Append("</rPr>");
            }
            sb.Append($"<t{ValueHasWhiteSpaces()}>");
            sb.Append(ConvertUtil.ExcelEscapeAndEncodeString(Text));
            sb.Append("</t>");
            sb.Append("</r>");
        }

        private void WriteRichTextColorAttributes(StringBuilder sb)
        {
            sb.Append("<color");
            if (ColorSettings.Auto)
            {
                sb.Append(" auto=\"1\"");
            }
            if (ColorSettings.Indexed != null)
            {
                sb.Append($" indexed=\"{ColorSettings.Indexed.Value}\"");
            }
            if (ColorSettings.Rgb != Color.Empty)
            {
                sb.Append($" rgb=\"{(ColorSettings.Rgb.ToArgb()).ToString("X").PadLeft(8, '0')}\"");
            }
            if (ColorSettings.Theme != null)
            {
                sb.Append($" theme=\"{(int)ColorSettings.Theme.Value}\"");
            }
            if (ColorSettings.Tint != null)
            {
                sb.Append($" tint=\"{ColorSettings.Tint.Value.ToString(CultureInfo.InvariantCulture)}\"");
            }
            sb.Append("/>");
        }

        bool HasDefaultValue
        {
            get
            {
                return       Bold == false &&
                           Italic == false &&
                           Strike == false &&
                    VerticalAlign == ExcelVerticalAlignmentFont.None &&
                             Size == 0 &&
                    String.IsNullOrEmpty(FontName) &&
                            Color == Color.Empty &&
                          Charset == null &&
                           Family == null &&
                    UnderLineType == null;
                        //Outline == false &&
                         //Shadow == false &&
                       //Condense == false &&
                         //Extend == false &&
                         //Scheme == null;
            }
        }
    }
}
