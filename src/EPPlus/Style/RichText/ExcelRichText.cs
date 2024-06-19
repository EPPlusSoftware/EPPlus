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
    /// A richtext part
    /// </summary>
    public class ExcelRichText
    {
        /// <summary>
        /// A referens to the richtext collection
        /// </summary>
        internal ExcelRichTextCollection _collection { get; set; }

        #region RichText Properties Attributes
        /// <summary>
        /// Preserves whitespace. Default true
        /// </summary>
        public bool PreserveSpace { get; set; } = true;

        /// <summary>
        /// Bold text
        /// </summary>
        public bool Bold { get; set; } = false;

        /// <summary>
        /// Italic text
        /// </summary>
        public bool Italic { get; set; } = false;

        /// <summary>
        /// Strike-out text
        /// </summary>
        public bool Strike { get; set; } = false;

        /// <summary>
        /// Underlined text
        /// <para/>
        /// True sets <see cref="UnderLineType">UnderLineType</see> to  <see cref="ExcelUnderLineType.Single">Single</see>
        /// <para/>
        /// False sets <see cref="UnderLineType">UnderLineType</see> to <see cref="ExcelUnderLineType.Single">None</see>
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return UnderLineType != ExcelUnderLineType.None;
            }
            set
            {
                UnderLineType = value ? ExcelUnderLineType.Single : ExcelUnderLineType.None;
            }
        }

        /// <summary>
        /// Vertical Alignment
        /// </summary>
        public ExcelVerticalAlignmentFont? VerticalAlign { get; set; } = ExcelVerticalAlignmentFont.None;

        /// <summary>
        /// Font size
        /// </summary>
        public float Size { get; set; } = 0;

        /// <summary>
        /// Name of the font
        /// </summary>
        public string FontName { get; set; } = "";

        /// <summary>
        /// Text color.
        /// Also see <seealso cref="ColorSettings"/>
        /// </summary>
        public Color Color
        {
            get
            {
                Color ret = Color.Empty;
                if (ColorSettings.Rgb != Color.Empty)
                {
                    ret = ColorSettings.Rgb;
                }
                else if (ColorSettings.Indexed.HasValue)
                {
                    ret = _collection._wb.Styles.GetIndexedColor(ColorSettings.Indexed.Value);
                }
                else if (ColorSettings.Theme.HasValue)
                {
                    ret = Utils.ColorConverter.GetThemeColor(_collection._wb.ThemeManager.GetOrCreateTheme(), ColorSettings.Theme.Value);
                }
                else if (ColorSettings.Auto == true)
                {
                    ret = Color.Black;
                }
                if (ColorSettings.Tint.HasValue)
                {
                    return Utils.ColorConverter.ApplyTint(ret, ColorSettings.Tint.Value);
                }
                return ret;
            }
            set
            {
                ColorSettings.Rgb = value;
            }
        }

        /// <summary>
        /// Color settings.
        /// <seealso cref="Color"/>
        /// </summary>
        public ExcelRichTextColor ColorSettings { get; set; }

        /// <summary>
        /// Characterset to use
        /// </summary>
        public int Charset { get; set; } = 0;

        /// <summary>
        /// Font family
        /// </summary>
        public int Family { get; set; } = 0;

        /// <summary>
        /// Underline type of text
        /// </summary>
        public ExcelUnderLineType UnderLineType { get; set; } = ExcelUnderLineType.None;

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
                _text = value;
            }
        }

        internal ExcelRichText(string text, ExcelRichTextCollection collection)
        {
            _collection = collection;
            ColorSettings = new ExcelRichTextColor();
            Text = text;
        }

        internal ExcelRichText(XmlReader xr, ExcelRichTextCollection collection)
        {
            ColorSettings = new ExcelRichTextColor();
            _collection = collection;
            string text = null;
            if (xr.LocalName == "rPr" && xr.NodeType == XmlNodeType.Element)
            {
                ReadrPr(xr);
                xr.Read();
            }
            if (xr.LocalName == "t" && xr.NodeType == XmlNodeType.Element)
            {
                text = xr.ReadElementContentAsString();
                Text = ConvertUtil.ExcelDecodeString(text);
            }
        }

        internal ExcelRichText(ExcelRichText rt, ExcelRichTextCollection collection)
        {
            _collection = collection;
            Text = rt.Text;
            Bold = rt.Bold;
            Italic = rt.Italic;
            PreserveSpace = rt.PreserveSpace;
            Strike = rt.Strike;
            UnderLine = rt.UnderLine;
            VerticalAlign = rt.VerticalAlign;
            Size = rt.Size;
            FontName = rt.FontName;
            ColorSettings = new ExcelRichTextColor();
            ColorSettings.Tint = rt.ColorSettings.Tint;
            ColorSettings.Indexed = rt.ColorSettings.Indexed;
            ColorSettings.Theme = rt.ColorSettings.Theme;
            ColorSettings.Rgb = rt.ColorSettings.Rgb;
            ColorSettings.Auto = rt.ColorSettings.Auto;
            Charset = rt.Charset;
            Family = rt.Family;
            UnderLineType = rt.UnderLineType;
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
                case "none":
                    return ExcelUnderLineType.None;
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
        internal void ReadrPr(XmlReader xr)
        {
            while (xr.Read())
            {
                if (xr.LocalName == "rPr") break;

                if(xr.NodeType == XmlNodeType.EndElement) 
                {
                    continue;
                }

                switch (xr.LocalName)
                {
                    case "b":
                        Bold = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "i":
                        Italic = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "strike":
                        Strike = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "u":
                        UnderLineType = GetUnderlineType(xr.GetAttribute("val"));
                        break;
                    case "vertAlign":
                        VerticalAlign = xr.GetAttribute("val").ToEnum<ExcelVerticalAlignmentFont>(ExcelVerticalAlignmentFont.None);
                        break;
                    case "sz":
                        if (ConvertUtil.TryParseNumericString(xr.GetAttribute("val"), out double num))
                        {
                            Size = Convert.ToSingle(num);
                        }
                        break;
                    case "rFont":
                        FontName = xr.GetAttribute("val");
                        break;
                    case "charset":
                        Charset = int.Parse(xr.GetAttribute("val"));
                        break;
                    case "family":
                        Family = int.Parse(xr.GetAttribute("val"));
                        break;
                    case "color":
                        ReadColor(xr);
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
        }

        internal void ReadColor(XmlReader xr)
        {
            if(ColorSettings == null)
            {
                ColorSettings = new ExcelRichTextColor();
            }
            int num;
            var auto = xr.GetAttribute("auto");
            if (int.TryParse(auto, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
            {
                ColorSettings.Auto = result > 0 || result < 0 ? true : false;
            }
            if (int.TryParse(xr.GetAttribute("indexed"), NumberStyles.Integer, CultureInfo.InvariantCulture, out num))
            {
                ColorSettings.Indexed = num;
            }
            var rgb = xr.GetAttribute("rgb");
            if (!String.IsNullOrEmpty(rgb))
            {
                ColorSettings.Rgb = ExcelDrawingRgbColor.GetColorFromString(rgb);
            }
            if (int.TryParse(xr.GetAttribute("theme"), NumberStyles.Integer, CultureInfo.InvariantCulture, out num))
            {
                ColorSettings.Theme = (eThemeSchemeColor)num;
            }
            var tint = xr.GetAttribute("tint");
            if (ConvertUtil.TryParseNumericString(tint, out double d, CultureInfo.InvariantCulture))
            {
                ColorSettings.Tint = d;
            }
        }

        /// <summary>
        /// Write RichTextAttributes
        /// </summary>
        /// <param name="sb"></param>
        internal void WriteRichTextAttributes(StringBuilder sb)
        {
            if (!string.IsNullOrEmpty(Text))
            {
                sb.Append("<r>");
                if (!HasDefaultValue)
                {
                    sb.Append("<rPr>");
                    if (!String.IsNullOrEmpty(FontName))
                    {
                        sb.Append($"<rFont val=\"{FontName}\"/>");
                    }
                    if (Charset != 0)
                    {
                        sb.Append($"<charset val=\"{Charset}\"/>");
                    }
                    if (Family != 0)
                    {
                        sb.Append($"<family val=\"{Family}\"/>");
                    }
                    if (Bold == true)
                    {
                        sb.Append($"<b/>");
                    }
                    if (Italic == true)
                    {
                        sb.Append($"<i/>");
                    }
                    if (Strike == true)
                    {
                        sb.Append($"<strike/>");
                    }
                    if (ColorSettings.HasAttributes)
                    {
                        WriteRichTextColorAttributes(sb);
                    }
                    if (Size > 0)
                    {
                        sb.Append($"<sz val=\"{Size.ToString(CultureInfo.InvariantCulture)}\"/>");
                    }
                    if (UnderLine)
                    {
                        sb.Append($"<u val=\"{UnderLineType.ToEnumString()}\"/>");
                    }
                    if (VerticalAlign != null && VerticalAlign != ExcelVerticalAlignmentFont.None)
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
        }

        internal void WriteRichTextColorAttributes(StringBuilder sb)
        {
            sb.Append("<color");
            if (ColorSettings.Auto == true)
            {
                sb.Append(" auto=\"1\"");
            }
            if (ColorSettings.Indexed != null)
            {
                sb.Append($" indexed=\"{ColorSettings.Indexed.Value}\"");
            }
            if (ColorSettings.Rgb != Color.Empty)
            {
                sb.Append($" rgb=\"{ColorSettings.Rgb.ToArgb().ToString("X").PadLeft(8, '0')}\"");
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
        /// <summary>
        /// Has default value
        /// </summary>
        public bool HasDefaultValue
        {
            get
            {
                return  Bold == false &&
                      Italic == false &&
                      Strike == false &&
               VerticalAlign == ExcelVerticalAlignmentFont.None &&
                        Size == 0 &&
             String.IsNullOrEmpty(FontName) &&
                       Color == Color.Empty &&
                     Charset == 0 &&
                      Family == 0 &&
               UnderLineType == 0;
                   //Outline == false &&
                    //Shadow == false &&
                  //Condense == false &&
                    //Extend == false &&
                    //Scheme == null;
            }
        }
    }
}
