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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System.Drawing;
using System.Reflection.Emit;
using System.Xml;
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A richtext part
    /// </summary>
    public class ExcelParagraphTextRun : XmlHelper
    {
        /// <summary>
        /// for measuring
        /// </summary>
        string _defaultFontName;
        IPictureRelationDocument _prd;
        XmlNode _rootNode;
        internal ExcelParagraphTextRun(IPictureRelationDocument prd, XmlNamespaceManager ns, XmlNode topNode) : base(ns, topNode)
        {
            if(topNode.LocalName=="r")
            {
                _rootNode = topNode;
            }
            _prd = prd;
        }
        internal void SetDefaultFontName(string defaultName)
        {
            _defaultFontName = defaultName;
        }

        internal string GetTextRunFontName()
        {
            if (string.IsNullOrEmpty(LatinFont))
            {
                if (string.IsNullOrEmpty(ComplexFont))
                {
                    return _defaultFontName;
                }
                else
                {
                    return ComplexFont;
                }
            }
            else
            {
                return LatinFont;
            }
        }

        //internal ExcelParagraphTextRun(TextCharacterAttributes attributes, bool bold, double baseline, eTextCapsType capitalization, bool italic, double kerning, double spacing, eStrikeType strike, double fontSize, eUnderLineType underLine, ExcelDrawingFill fill, ExcelDrawingFill fill, Color color, Color underLineColor, string latinFont, string eastAsianFont, string complexFont, string symbolFont, bool rtl, string t)
        //{
        //    Attributes = attributes;
        //    Bold = bold;
        //    Baseline = baseline;
        //    Capitalization = capitalization;
        //    Italic = italic;
        //    Kerning = kerning;
        //    Spacing = spacing;
        //    Strike = strike;
        //    FontSize = fontSize;
        //    UnderLine = underLine;
        //    Fill = fill;
        //    Color = color;
        //    UnderLineColor = underLineColor;
        //    LatinFont = latinFont;
        //    EastAsianFont = eastAsianFont;
        //    ComplexFont = complexFont;
        //    SymbolFont = symbolFont;
        //    this.rtl = rtl;
        //    this.t = t;
        //}

  

        #region LineProperties
        //TODO: Line Properties
        #endregion LineProperties

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
                    _fill = new ExcelDrawingFill(_prd, NameSpaceManager, TopNode, "a:r", SchemaNodeOrder);
                }
                return _fill;
            }
        }

        //Below is quick-access to the drawing fill
        string _colorPath = "a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Sets the default color of the text.
        /// This sets the Fill to a SolidFill with the specified color.
        /// <remark>
        /// Use the Fill property for more options
        /// </remark>
        /// </summary>
        public Color Color
        {
            get
            {
                string col = GetXmlNodeString(_colorPath);
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
                Fill.Style = eFillStyle.SolidFill;
                Fill.SolidFill.Color.SetRgbColor(value);
            }
        }
        #endregion Basic fill

        //UnderlineLine underlineFill etc.
        #region Underline
        string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";
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

        string _fontLatinPath = "a:latin/@typeface";
        /// <summary>
        /// The latin typeface name
        /// </summary>
        public string LatinFont
        {
            get
            {
                return GetXmlNodeString(_fontLatinPath);
            }
            set
            {
                SetXmlNodeString(_fontLatinPath, value);
            }
        }
        string _fontEaPath = "a:ea/@typeface";
        /// <summary>
        /// The East Asian typeface name
        /// </summary>
        public string EastAsianFont
        {
            get
            {
                return GetXmlNodeString(_fontEaPath);
            }
            set
            {
                SetXmlNodeString(_fontEaPath, value);
            }
        }
        string _fontCsPath = "a:cs/@typeface";
        /// <summary>
        /// The complex font typeface name
        /// </summary>
        public string ComplexFont
        {
            get
            {
                return GetXmlNodeString(_fontCsPath);
            }
            set
            {
                SetXmlNodeString(_fontCsPath, value);
            }
        }

        string _fontSymPath = "a:sym/@typeface";
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

        #endregion FontNodes

        #region HyperLink
        #endregion Hyperlink

        //TODO:
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
        internal string Text;
        /// <summary>
        /// Creates the top nodes of the collection
        /// </summary>
    }
}
