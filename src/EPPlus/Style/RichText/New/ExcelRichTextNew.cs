﻿/*************************************************************************************************
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
using System.Text;
using System.Xml;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils;
using System.Runtime.InteropServices;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// A richtext part
    /// </summary>
    public class ExcelRichTextNew
    {
        internal ExcelRichTextNew(ExcelRichTextCollectionNew collection)
        {
            _collection = collection;
            ColorSettings = new ExcelRichTextColorNew(this);
        }

        /// <summary>
        /// The text
        /// </summary>
        public string Text { get; set; }
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

        internal void ReadrPr(XmlReader xr)
        {
            while(xr.Read())
            {
                if (xr.LocalName == "rPr") return;

                switch (xr.LocalName)
                {
                    case "rFont":
                        FontName = xr.GetAttribute("val");
                        break;
                    case "charset":
                        Charset = int.Parse(xr.GetAttribute("val"));
                        break;
                    case "family":
                        Family = int.Parse(xr.GetAttribute("val"));
                        break;
                    case "b":
                        Bold = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "i":
                        Italic = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
                        break;
                    case "strike":
                        Strike = ConvertUtil.ToBooleanString(xr.GetAttribute("val"), true);
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
                    case "color":
                        ColorSettings = new ExcelRichTextColorNew(xr, this);
                        break;
                    case "sz":
                        if(ConvertUtil.TryParseNumericString(xr.GetAttribute("val"),out double num)) 
                        {
                            Size = Convert.ToSingle( num );
                        }
                        break;
                    case "u":
                        UnderLineType = GetUnderlineType(xr.GetAttribute("val"));
                        //UnderLine = UnderLineType != ExcelUnderLineType.None;
                        break;
                    case "vertAlign":
                        VerticalAlign = xr.GetAttribute("val").ToEnum<ExcelVerticalAlignmentFont>(ExcelVerticalAlignmentFont.None);
                        break;
                    //case "scheme":
                    //    Scheme = xr.GetAttribute("val").ToEnum<eThemeFontCollectionType>(eThemeFontCollectionType.None);
                    //    break;
                }
            }
        }

        private ExcelUnderLineType GetUnderlineType(string v)
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

        internal void GetXML(StringBuilder sb)
        {
            sb.Append("<r>");
            if (!HasDefaultValue) 
            {
                sb.Append("<rPr>");
                if(!String.IsNullOrEmpty(FontName))
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
                if (Bold)
                {
                    sb.Append($"<b/>");
                }
                if (Italic)
                {
                    sb.Append($"<i/>");
                }
                if (Strike)
                {
                    sb.Append($"<strike/>");
                }
                if (Color != Color.Empty)
                {
                    ColorSettings.AppendXml(sb);
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
            sb.Append(Text);
            sb.Append("</t>");
            sb.Append("</r>");
        }

        string ValueHasWhiteSpaces()
        {
            if(Text != null && Text.Length > 0)
            {
                if(char.IsWhiteSpace(Text[0]) || char.IsWhiteSpace(Text[Text.Length - 1]))
                {
                    return " xml:space=\"preserve\"";
                }
            }
            return "";
        }

        bool HasDefaultValue { 
            get {
                return Bold == false &&
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
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return UnderLineType != null && UnderLineType != ExcelUnderLineType.None;
            }
            set
            {
                UnderLineType = value ? ExcelUnderLineType.Single : null;
            }
        }

        /// <summary>
        /// Vertical Alignment
        /// </summary>
        public ExcelVerticalAlignmentFont VerticalAlign { get; set; } = ExcelVerticalAlignmentFont.None;

        /// <summary>
        /// Font size
        /// </summary>
        public float Size { get; set; } = 0f;

        /// <summary>
        /// Name of the font
        /// </summary>
        public string FontName { get; set; } = string.Empty;


        /// <summary>
        /// Text color.
        /// Also see <seealso cref="ColorSettings"/>
        /// </summary>
        public Color Color
        {
            get
            {
                return ColorSettings.Color;
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
        public ExcelRichTextColorNew ColorSettings { get; set; }

        /// <summary>
        /// A referens to the richtext collection
        /// </summary>
        public ExcelRichTextCollectionNew _collection { get; set; }

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
}
