/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/15/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.EnumUtils;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A richtext part
    /// </summary>
    public abstract class ExcelParagraphTextRunBase : XmlHelper
    {
        /// <summary>
        /// for measuring
        /// </summary>
        string _defaultFontName;
        double _defaultFontSize;
        ExcelTextFont _dtr;
        internal IPictureRelationDocument _prd;
        ExcelDrawingParagraph _paragraph;

        internal ExcelParagraphTextRunBase(ExcelDrawingParagraph paragraph, XmlNamespaceManager ns, XmlNode topNode) : base(ns, topNode)
        {
            SchemaNodeOrder = ["rPr", "pPr", "t"];
            _paragraph = paragraph;
            _prd = _paragraph._prd;
            _dtr = _paragraph._paragraphs.FirstOrDefault()?.DefaultRunProperties;
        }

        internal List<string> SplitIntoLines()
        {
            var strLst = Text.Split("\r\n".ToArray()).ToList();
            return strLst;
        }

        internal ExcelDrawingParagraph Paragraph { get => _paragraph; }
        /// <summary>
        /// The type of text run
        /// </summary>
        public abstract eParagraphRunType Type { get; }

        internal MeasurementFont GetMeasurementFont()
        {
            var mf = new MeasurementFont()
            {
                FontFamily = string.IsNullOrEmpty(LatinFont) ? ComplexFont : LatinFont,
                Size = FontSize,
                Style = GetFontStyle()
            };

            if (string.IsNullOrEmpty(mf.FontFamily) || mf.Size <= 0 || float.IsNaN(mf.Size))
            {
                var defaultMeasurementFont = _paragraph.DefaultRunProperties.GetMeasureFont();

                if (string.IsNullOrEmpty(mf.FontFamily))
                {
                    mf.FontFamily = defaultMeasurementFont.FontFamily;
                }

                if(mf.Size <= 0 || float.IsNaN(mf.Size))
                {
                    mf.Size = defaultMeasurementFont.Size;
                }
            }

            return mf;
        }

        #region Basic Fill
        ExcelDrawingFill _fill;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_prd, NameSpaceManager, TopNode, "a:rPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }

        ////Below is quick-access to the drawing fill
        //string _colorPath = "a:rPr/a:solidFill/a:srgbClr/@val";
        ///// <summary>
        ///// Sets the default color of the text.
        ///// This sets the Fill to a SolidFill with the specified color.
        ///// <remark>
        ///// Use the Fill property for more options
        ///// </remark>
        ///// </summary>
        //public Color Color
        //{
        //    get
        //    {
        //        string col = GetXmlNodeString(_colorPath);
        //        if (col == "")
        //        {
        //            return Color.Empty;
        //        }
        //        else
        //        {
        //            return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
        //        }
        //    }
        //    set
        //    {
        //        Fill.Style = eFillStyle.SolidFill;
        //        Fill.SolidFill.Color.SetRgbColor(value);
        //    }
        //}
        #endregion Basic fill

        //UnderlineLine underlineFill etc.
        #region Underline
        string _underLineColorPath = "a:rPr/a:uFill/a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// The fonts underline color
        /// </summary>
        public Color UnderLineColor
        {
            get
            {
                string col = GetXmlNodeString(_underLineColorPath);
                if (col == "")
                {
                    return Color.Empty;
                }
                else
                {
                    return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
                }
            }
            set
            {
                SetXmlNodeString(_underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
            }
        }

        #endregion Underline

        #region FontNodes

        string _fontLatinPath = "a:rPr/a:latin/@typeface";
        /// <summary>
        /// The latin typeface name
        /// </summary>
        public string LatinFont
        {
            get
            {
                var v=GetXmlNodeString(_fontLatinPath);
                if(string.IsNullOrEmpty(v))
                {
                    v = _dtr.LatinFont;
                }
                return v;
            }
            set
            {
                SetXmlNodeString(_fontLatinPath, value);
            }
        }
        string _fontEaPath = "a:rPr/a:ea/@typeface";
        /// <summary>
        /// The East Asian typeface name
        /// </summary>
        public string EastAsianFont
        {
            get
            {
                var v = GetXmlNodeString(_fontEaPath);
                if (string.IsNullOrEmpty(v))
                {
                    v = _dtr.EastAsianFont;
                }
                return v;
            }
            set
            {
                SetXmlNodeString(_fontEaPath, value);
            }
        }
        string _fontCsPath = "a:rPr/a:cs/@typeface";
        /// <summary>
        /// The complex font typeface name
        /// </summary>
        public string ComplexFont
        {
            get
            {
                var v = GetXmlNodeString(_fontCsPath);
                if (string.IsNullOrEmpty(v))
                {
                    v = _dtr.ComplexFont;
                }
                return v;
            }
            set
            {
                SetXmlNodeString(_fontCsPath, value);
            }
        }

        string _fontSymPath = "a:rPr/a:sym/@typeface";
        /// <summary>
        /// The symbol font typeface name
        /// </summary>
        public string SymbolFont
        {
            get
            {
                return GetXmlNodeString(_fontSymPath);
            }
            set
            {
                SetXmlNodeString(_fontSymPath, value);
            }
        }
        string _boldPath = "a:rPr/@b";
        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool FontBold
        {
            get
            {
                var v=GetXmlNodeBoolNullable(_boldPath);
                if(v==null)
                {
                    v = _dtr.Bold;
                }
                return v??false;
            }
            set
            {
                SetXmlNodeString(_boldPath, value ? "1" : "0");
            }
        }
        string _underLinePath = "a:rPr/@u";
        /// <summary>
        /// The fonts underline style
        /// </summary>
        public eUnderLineType FontUnderLine
        {
            get
            {
                var v = GetXmlNodeString(_underLinePath)?.TranslateUnderline();
                if (v == null)
                {
                    v = _dtr.UnderLine;
                }
                return v ?? eUnderLineType.None;
            }
            set
            {
                SetXmlNodeString(_underLinePath, value.TranslateUnderlineText());
            }
        }

        string _italicPath = "a:rPr/@i";
        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool FontItalic
        {
            get
            {
                var v = GetXmlNodeBoolNullable(_italicPath);
                if (v == null)
                {
                    v = _dtr.Italic;
                }
                return v ?? false;

            }
            set
            {
                SetXmlNodeString(_italicPath, value ? "1" : "0");
            }
        }
        string _strikePath = "a:rPr/@strike";
        /// <summary>
        /// Font strike out type
        /// </summary>
        public eStrikeType FontStrike
        {
            get
            {
                var v = GetXmlNodeString(_strikePath)?.TranslateStrikeType();
                if (v == null)
                {
                    v = _dtr.Strike;
                }
                return v ?? eStrikeType.No;
            }
            set
            {
                SetXmlNodeString(_strikePath, value.TranslateStrikeTypeText());
            }
        }
        string _sizePath = "a:rPr/@sz";
        /// <summary>
        /// Font size
        /// </summary>
        public float FontSize
        {
            get
            {
                var v = GetXmlNodeDoubleNull(_sizePath);
                if (v == null)
                {
                    return _dtr.Size;
                }
                else
                {
                    return (float)(v / 100);
                }
            }
            set
            {
                SetXmlNodeString(_sizePath, ((int)(value * 100)).ToString());
            }
        }
        string _kernPath = "a:rPr/@kern";
        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run
        /// </summary>
        public double Kerning
        {
            get
            {
                var v = GetXmlNodeDoubleNull(_kernPath);
                if(v==null)
                {
                    return _dtr.Kerning;
                }
                else
                {
                    return (float)(v / 100);
                }
            }
            set
            {
                SetXmlNodeFontSize(_kernPath, value, "Kerning");
            }
        }
        string _capPath = "a:rPr/@cap";
        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public eTextCapsType Capitalization
        {
            get
            {
                var v = GetXmlNodeString(_capPath);
                if(v==null)
                {
                    return _dtr.Capitalization;
                }
                else
                {
                    switch (v)
                    {
                        case "all":
                            return eTextCapsType.All;
                        case "small":
                            return eTextCapsType.Small;
                        default:
                            return eTextCapsType.None;
                    }
                }
            }
            set
            {                
                SetXmlNodeString(_capPath, value.ToEnumString());
            }
        }

        string _baselinePath = "a:rPr/@baseline";
        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public double Baseline
        {
            get
            {
                var v=GetXmlNodeDoubleNull(_baselinePath);
                if (v == null)
                {
                        return _dtr.Baseline;
                }
                else
                {
                    return v.Value;
                }
            }
            set
            {
                SetXmlNodePercentage(_baselinePath, value);
            }
        }
        string _highlightPath = "a:rPr/a:highlight";
        /// <summary>
        /// The highlight color.
        /// </summary>
        public ExcelDrawingColorManager HighlightColor        
        {
            get
            {
                return new ExcelDrawingColorManager(NameSpaceManager, TopNode, _highlightPath, SchemaNodeOrder);
            }
        }
        string _spacingPath = "a:rPr/@spc";
        public double Spacing
        {
            get
            {
                var v = GetXmlNodeDoubleNull(_spacingPath);
                if(v==null)
                {
                    return _dtr.Spacing;
                }
                else
                {
                    return v.Value / 100;
                }
            }
            set
            {
                SetXmlNodeDouble(_spacingPath, value * 100);
            }
        }
        /// <summary>
        /// Set the font style properties
        /// </summary>
        /// <param name="name">Font family name</param>
        /// <param name="size">Font size</param>
        /// <param name="bold"></param>
        /// <param name="italic"></param>
        /// <param name="underline"></param>
        /// <param name="strikeout"></param>
        public void SetFromFont(string name, float size, bool bold = false, bool italic = false, bool underline = false, bool strikeout = false)
        {
            LatinFont = name;
            ComplexFont = name;
            FontSize = size;
            if (bold) FontBold = bold;
            if (italic) FontItalic = italic;
            if (underline) FontUnderLine = eUnderLineType.Single;
            if (strikeout) FontStrike = eStrikeType.Single;
        }

        //internal void GetHeightInPixels(out float textWidth, out float textHeight, string text)
        //{
        //    var tm = _prd.Package.Settings.TextSettings.PrimaryTextMeasurer;
        //    _prd.Package.Workbook.Styles.GetNormalStyle();
        //    MeasurementFont f = GetMeasureFont();
        //    var b = tm.MeasureText(text, f);
        //    textWidth = b.Width;
        //    textHeight = b.Height;
        //}
        internal MeasurementFont GetMeasureFont()
        {
            var lf = LatinFont;
            if (string.IsNullOrEmpty(lf))
            {
                lf = _dtr.LatinFont;
            }
            var cf = ComplexFont;
            if (string.IsNullOrEmpty(cf))
            {
                cf = _paragraph.DefaultRunProperties.ComplexFont;
            }
            var eaf = ComplexFont;
            if (string.IsNullOrEmpty(eaf))
            {
                eaf = _paragraph.DefaultRunProperties.EastAsianFont;
            }
            if (double.IsNaN(FontSize))
            {
                FontSize = _paragraph.DefaultRunProperties.Size;
            }
            return FontUtil.GetMeasureFont(lf, cf, eaf, FontSize, GetFontStyle(), _prd.Package);
        }
        private MeasurementFontStyles GetFontStyle()
        {
            MeasurementFontStyles ret = MeasurementFontStyles.Regular;
            if (FontBold)
            {
                ret |= MeasurementFontStyles.Bold;
            }
            if (FontItalic)
            {
                ret |= MeasurementFontStyles.Italic;
            }
            if (FontUnderLine != eUnderLineType.None)
            {
                ret |= MeasurementFontStyles.Underline;
            }
            return ret;
        }

        internal bool IsEmpty
        {
            get
            {
                return TopNode == null || (TopNode.ChildNodes.Count == 0 && TopNode.Attributes.Count == 0);
            }
        }

        #endregion FontNodes

        #region HyperLink
        #endregion Hyperlink

        /// <summary>
        /// Right to left
        /// If ommitted it returns false AKA (left-to-right)
        /// </summary>
        internal bool rtl;

        //TODO:
        #region ExtLst-OfficeArtExtensionList
        #endregion ExtLst-OfficeArtExtensionList

        /// <summary>
        /// Actual text for the text run
        /// </summary>
        public abstract string Text
        {
            get;
            set;
        }
        /// <summary>
        /// If the text item is the first item in the paragraph
        /// </summary>
        public bool IsFirstInParagraph
        {
            get
            {
                return _paragraph.TextRuns.IndexOf(this) == 0;
            }
        }
        /// <summary>
        /// If the text item is the last item in the paragraph
        /// </summary>
        public bool IsLastInParagraph
        {
            get
            {
                return _paragraph.TextRuns.IndexOf(this) == _paragraph.TextRuns.Count-1;
            }
        }
    }
}
