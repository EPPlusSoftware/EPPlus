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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public class ExcelTextFont : XmlHelper
    {
        string _path;
        internal XmlNode _rootNode;
        Action _initXml;
        IPictureRelationDocument _pictureRelationDocument;
        internal ExcelTextFont(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder, Action initXml=null)
            : base(namespaceManager, rootNode)
        {
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "bodyPr", "lstStyle","p", "pPr", "defRPr", "solidFill","highlight", "uFill", "latin","ea", "cs","sym","hlinkClick","hlinkMouseOver","rtl", "r", "rPr", "t" });
            _rootNode = rootNode;
            _pictureRelationDocument = pictureRelationDocument;
            _initXml = initXml;
            if (path != "")
            {
                XmlNode node = rootNode.SelectSingleNode(path, namespaceManager);
                if (node != null)
                {
                    TopNode = node;
                }
            }
            _path = path;

            textRun = new RegularTextRun();
            //Reads attribute values into textRun
            ParseAttributesFromXML();
        }

        RegularTextRun textRun;

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
                    _fill = new ExcelDrawingFill(_pictureRelationDocument, NameSpaceManager, _rootNode, _path, SchemaNodeOrder, CreateTopNode);
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
                CreateTopNode();
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
                CreateTopNode();
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
                CreateTopNode();
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
                CreateTopNode();
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
                CreateTopNode();
                SetXmlNodeString(_fontSymPath, value);
            }
        }

        #endregion FontNodes

        #region HyperLink
        #endregion Hyperlink

        string rtlPath = "/a:rtl/@w:val";

        /// <summary>
        /// Right to left
        /// If ommitted it returns false AKA (left-to-right)
        /// </summary>
        internal bool rtl
        {
            get
            {
                return GetBoolFromNullString(rtlPath);
            }
            set
            {
                SetBoolNode(rtlPath, value);
            }
        }

        #region ExtLst-OfficeArtExtensionList
        #endregion ExtLst-OfficeArtExtensionList


        /// <summary> 
        /// Creates the top nodes of the collection
        /// </summary>
        protected internal void CreateTopNode()
        {
            if (_path!="" && TopNode==_rootNode)
            {
                _initXml?.Invoke();
                if (TopNode == _rootNode && string.IsNullOrEmpty(_path)==false)
                {
                    CreateNode(_path);
                    TopNode = _rootNode.SelectSingleNode(_path, NameSpaceManager);
                    CreateNode("../../../a:bodyPr");
                    CreateNode("../../../a:lstStyle");
                }
            }
            else if (TopNode.ParentNode?.ParentNode?.ParentNode?.LocalName == "rich")
            {
                CreateNode("../../../a:bodyPr");
                CreateNode("../../../a:lstStyle");
            }
        }

        string _boldPath = "@b";

        internal void ParseAttributesFromXML()
        {
            textRun.Bold = GetXmlNodeBool(_boldPath);
            textRun.UnderLine = GetXmlNodeString(_underLinePath).TranslateUnderline();
            textRun.Italic = GetXmlNodeBool(_italicPath);
            textRun.Strike = GetXmlNodeString(_strikePath).TranslateStrikeType();

            //TODO: handle Int.MIN same as before
            textRun.FontSize = GetXmlNodeInt(_sizePath);
            textRun.Kerning = GetXmlNodeFontSize(_kernPath);

            textRun.Capitalization = GetXmlNodeString($"{_path}/@cap").ToEnum(eTextCapsType.None);
            textRun.Baseline = GetXmlNodePercentage($"{_path}/@baseline") ?? 0;
        }
        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return textRun.Bold;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_boldPath, value ? "1" : "0");
                textRun.Bold = value;
            }
        }
        string _underLinePath = "@u";
        /// <summary>
        /// The fonts underline style
        /// </summary>
        public eUnderLineType UnderLine
        {
            get
            {
                return textRun.UnderLine;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_underLinePath, value.TranslateUnderlineText());
                textRun.UnderLine = value;
            }
        }

        internal void SetFromXml(XmlElement copyFromElement)
        {
            CreateTopNode();
            foreach (XmlAttribute a in copyFromElement.Attributes)
            {
                ((XmlElement)TopNode).SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
            if(copyFromElement.HasChildNodes && !TopNode.HasChildNodes)
            {
                TopNode.InnerXml = copyFromElement.InnerXml;
            }
        }
        string _italicPath = "@i";
        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return textRun.Italic;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_italicPath, value ? "1" : "0");
                textRun.Italic = value;
            }
        }
        string _strikePath = "@strike";
        /// <summary>
        /// Font strike out type
        /// </summary>
        public eStrikeType Strike
        {
            get
            {
                return textRun.Strike;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_strikePath, value.TranslateStrikeTypeText());
                textRun.Strike = value;
            }
        }
        string _sizePath = "@sz";
        /// <summary>
        /// Font size
        /// </summary>
        public float Size
        {
            get
            {
                return (float)textRun.FontSize;
                //var c = GetXmlNodeInt(_sizePath);
                //if(c==int.MinValue)
                //{
                //    return c;
                //}
                //return c / 100;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_sizePath, ((int)(value * 100)).ToString());
                textRun.FontSize = (float)value;
            }
        }
        string _kernPath = "@kern";
        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run
        /// </summary>
        public double Kerning
        {
            get
            {
                return textRun.Kerning;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeFontSize(_kernPath, value, "Kerning");
                textRun.Kerning = value;
            }
        }
        ///// <summary>
        ///// The baseline for both the superscript and subscript fonts in percentage
        ///// </summary>
        //public double Baseline { get => textRun.Attributes.Baseline; set => textRun.Attributes.Baseline = value; }
        ///// <summary>
        ///// The capitalization that is to be applied
        ///// </summary>
        //public eTextCapsType Capitalization { get => textRun.Attributes.Capitalization; set => textRun.Attributes.Capitalization = value; }

        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public eTextCapsType Capitalization
        {
            get
            {
                return textRun.Capitalization;
            }
            set
            {
                SetXmlNodeString($"{_path}/@kern", value.ToEnumString());
                textRun.Capitalization = value;
            }
        }

        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public double Baseline
        {
            get
            {
                return textRun.Baseline;
            }
            set
            {
                SetXmlNodePercentage($"{_path}/@baseline", value);
                textRun.Baseline = value;
            }
        }

        #region Methods
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
            Size = size;
            if (bold) Bold = bold;
            if (italic) Italic = italic;
            if (underline) UnderLine = eUnderLineType.Single;
            if (strikeout) Strike = eStrikeType.Single;            
        }

        internal void GetHeightInPixels(out float textWidth, out float textHeight, string text)
        {
            var tm = _pictureRelationDocument.Package.Settings.TextSettings.PrimaryTextMeasurer;
            _pictureRelationDocument.Package.Workbook.Styles.GetNormalStyle();
            MeasurementFont f = GetMeasureFont();
            var b = tm.MeasureText(text, f);
            textWidth = b.Width;
            textHeight = b.Height;
        }

        internal MeasurementFont GetMeasureFont()
        {
            return new MeasurementFont()
            {
                FontFamily = LatinFont,
                Size = Size,
                Style = GetFontStyle()
            };
        }

        private MeasurementFontStyles GetFontStyle()
        {
            MeasurementFontStyles ret = MeasurementFontStyles.Regular;
            if (Bold)
            {
                ret |= MeasurementFontStyles.Bold;
            }
            if (Italic)
            {
                ret |= MeasurementFontStyles.Italic;
            }
            if (UnderLine != eUnderLineType.None)
            {
                ret |= MeasurementFontStyles.Underline;
            }
            return ret;
        }
        internal bool IsEmpty
        {
            get
            {
                return TopNode==null || (TopNode.ChildNodes.Count==0 && TopNode.Attributes.Count==0);               
            }            
        }

        

        internal XmlElement PathElement
        {
            get
            {
                var node = (XmlElement)GetNode(_path);
                if (node == null)
                {
                    return (XmlElement)CreateNode(_path);
                }
                else
                {
                    return node;
                }
            }
        }
        #endregion Methods
    }
}
