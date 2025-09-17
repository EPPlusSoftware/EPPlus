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
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Drawing;
using System.Reflection.Emit;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public class ExcelTextFontXml : ExcelTextFont
    {
        XmlHelper _xml;
        string _path;
        internal XmlNode _rootNode;
        Action _initXml;
        IPictureRelationDocument _pictureRelationDocument;
        internal ExcelTextFontXml(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder, Action initXml = null)            
        {
            _xml = XmlHelperFactory.Create(namespaceManager, rootNode);
            _xml.AddSchemaNodeOrder(schemaNodeOrder, new string[] { "bodyPr", "lstStyle", "p", "pPr", "defRPr", "solidFill", "highlight", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "r", "rPr", "t" });
            _rootNode = rootNode;
            _pictureRelationDocument = pictureRelationDocument;
            _initXml = initXml;
            if (path != "")
            {
                XmlNode node = rootNode.SelectSingleNode(path, namespaceManager);
                if (node != null)
                {
                    _xml.TopNode = node;
                }
            }
            _path = path;
        }
        internal XmlHelper XmlHelper { get { return _xml; } }
        string _fontLatinPath = "a:latin/@typeface";
        /// <summary>
        /// The latin typeface name
        /// </summary>
        public override string LatinFont
        {
            get
            {
                return _xml.GetXmlNodeString(_fontLatinPath);
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_fontLatinPath, value);
            }
        }
        string _fontEaPath = "a:ea/@typeface";
        /// <summary>
        /// The East Asian typeface name
        /// </summary>
        public override string EastAsianFont
        {
            get
            {
                return _xml.GetXmlNodeString(_fontEaPath);
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_fontEaPath, value);
            }
        }
        string _fontCsPath = "a:cs/@typeface";
        /// <summary>
        /// The complex font typeface name
        /// </summary>
        public override string ComplexFont
        {
            get
            {
                return _xml.GetXmlNodeString(_fontCsPath);
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_fontCsPath, value);
            }
        }

        /// <summary>
        /// Creates the top nodes of the collection
        /// </summary>
        protected internal void CreateTopNode()
        {
            if (_path != "" && _xml.TopNode == _rootNode)
            {
                _initXml?.Invoke();
                if (_xml.TopNode == _rootNode && string.IsNullOrEmpty(_path) == false)
                {
                    _xml.CreateNode(_path);
                    _xml.TopNode = _rootNode.SelectSingleNode(_path, _xml.NameSpaceManager);
                    _xml.CreateNode("../../../a:bodyPr");
                    _xml.CreateNode("../../../a:lstStyle");
                }
            }
            else if (_xml.TopNode.ParentNode?.ParentNode?.ParentNode?.LocalName == "rich")
            {
                _xml.CreateNode("../../../a:bodyPr");
                _xml.CreateNode("../../../a:lstStyle");
            }
        }
        string _boldPath = "@b";
        /// <summary>
        /// If the font is bold
        /// </summary>
        public override bool Bold
        {
            get
            {
                return _xml.GetXmlNodeBool(_boldPath);
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_boldPath, value ? "1" : "0");
            }
        }
        string _underLinePath = "@u";
        /// <summary>
        /// The fonts underline style
        /// </summary>
        public override eUnderLineType UnderLine
        {
            get
            {
                return _xml.GetXmlNodeString(_underLinePath).TranslateUnderline();
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_underLinePath, value.TranslateUnderlineText());
            }
        }

        internal void SetFromXml(XmlElement copyFromElement)
        {
            CreateTopNode();
            foreach (XmlAttribute a in copyFromElement.Attributes)
            {
                ((XmlElement)_xml.TopNode).SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
            if (copyFromElement.HasChildNodes && !_xml.TopNode.HasChildNodes)
            {
                _xml.TopNode.InnerXml = copyFromElement.InnerXml;
            }
        }

        string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// The fonts underline color
        /// </summary>
        public override Color UnderLineColor
        {
            get
            {
                string col = _xml.GetXmlNodeString(_underLineColorPath);
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
                _xml.SetXmlNodeString(_underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
            }
        }
        string _italicPath = "@i";
        /// <summary>
        /// If the font is italic
        /// </summary>
        public override bool Italic
        {
            get
            {
                return _xml.GetXmlNodeBool(_italicPath);
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_italicPath, value ? "1" : "0");
            }
        }
        string _strikePath = "@strike";
        /// <summary>
        /// Font strike out type
        /// </summary>
        public override eStrikeType Strike
        {
            get
            {
                return _xml.GetXmlNodeString(_strikePath).TranslateStrikeType();
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_strikePath, value.TranslateStrikeTypeText());
            }
        }
        string _sizePath = "@sz";
        /// <summary>
        /// Font size
        /// </summary>
        public override float Size
        {
            get
            {
                var c = _xml.GetXmlNodeInt(_sizePath);
                if (c == int.MinValue)
                {
                    return c;
                }
                return c / 100;
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString(_sizePath, ((int)(value * 100)).ToString());
            }
        }
        ExcelDrawingFill _fill;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public override ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_pictureRelationDocument, _xml.NameSpaceManager, _rootNode, _path, _xml.SchemaNodeOrder, CreateTopNode);
                }
                return _fill;
            }
        }
        string _colorPath = "a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Sets the default color of the text.
        /// This sets the Fill to a SolidFill with the specified color.
        /// <remark>
        /// Use the Fill property for more options
        /// </remark>
        /// </summary>
        public override Color Color
        {
            get
            {
                string col = _xml.GetXmlNodeString(_colorPath);
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
        string _kernPath = "@kern";
        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run
        /// </summary>
        public override double Kerning
        {
            get
            {
                return _xml.GetXmlNodeFontSize(_kernPath);
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeFontSize(_kernPath, value, "Kerning");
            }
        }
        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public override eTextCapsType Capitalization
        {
            get
            {
                switch (_xml.GetXmlNodeString($"{_path}/@cap"))
                {
                    case "all":
                        return eTextCapsType.All;
                    case "small":
                        return eTextCapsType.Small;
                    default:
                        return eTextCapsType.None;
                }
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodeString($"{_path}/@cap", value.ToEnumString());
            }
        }

        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public override double Baseline
        {
            get
            {
                return _xml.GetXmlNodeDouble($"{_path}/@baseline");
            }
            set
            {
                CreateTopNode();
                _xml.SetXmlNodePercentage($"{_path}/@baseline", value);
            }
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
        internal override bool IsEmpty
        {
            get
            {
                return _xml.TopNode == null || (_rootNode == _xml.TopNode && string.IsNullOrEmpty(_path) == false) || (_xml.TopNode.ChildNodes.Count == 0 && _xml.TopNode.Attributes.Count == 0);
            }
        }
        internal override XmlElement PathElement
        {
            get
            {
                var node = (XmlElement)_xml.GetNode(_path);
                if (node == null)
                {
                    return (XmlElement)_xml.CreateNode(_path);
                }
                else
                {
                    return node;
                }
            }
        }

    }
    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public class ExcelTextFontRichText : ExcelTextFont
    {
        internal ExcelParagraphTextRunBase _textRun;
        internal ExcelTextFontRichText(ExcelParagraphTextRunBase textRun)
        {
            _textRun = textRun;
        }
        /// <summary>
        /// The latin typeface name
        /// </summary>
        public override string LatinFont
        {
            get
            {
                return _textRun.LatinFont;
            }
            set
            {
                _textRun.LatinFont = value;
            }
        }
        /// <summary>
        /// The East Asian typeface name
        /// </summary>
        public override string EastAsianFont
        {
            get
            {
                return _textRun.EastAsianFont;
            }
            set
            {
                _textRun.EastAsianFont = value;
            }
        }
        /// <summary>
        /// The complex font typeface name
        /// </summary>
        public override string ComplexFont
        {
            get
            {
                return _textRun.ComplexFont;
            }
            set
            {
                _textRun.ComplexFont = value;
            }
        }

        /// <summary>
        /// If the font is bold
        /// </summary>
        public override bool Bold
        {
            get
            {
                return _textRun.FontBold;
            }
            set
            {
                _textRun.FontBold = value;
            }
        }
        /// <summary>
        /// The fonts underline style
        /// </summary>
        public override eUnderLineType UnderLine
        {
            get
            {
                return _textRun.FontUnderLine;
            }
            set
            {
                _textRun.FontUnderLine = value;
            }
        }

        /// <summary>
        /// The fonts underline color
        /// </summary>
        public override Color UnderLineColor
        {
            get
            {
                return _textRun.UnderLineColor;
            }
            set
            {
                _textRun.UnderLineColor = value;
            }
        }
        /// <summary>
        /// If the font is italic
        /// </summary>
        public override bool Italic
        {
            get
            {
                return _textRun.FontItalic;
            }
            set
            {
                _textRun.FontItalic = value;
            }
        }
        /// <summary>
        /// Font strike out type
        /// </summary>
        public override eStrikeType Strike
        {
            get
            {
                return _textRun.FontStrike;
            }
            set
            {
                _textRun.FontStrike = value;
            }
        }
        /// <summary>
        /// Font size
        /// </summary>
        public override float Size
        {
            get
            {
                return _textRun.FontSize;
            }
            set
            {
                _textRun.FontSize = value;
            }
        }
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public override ExcelDrawingFill Fill
        {
            get
            {
                return _textRun.Fill;
            }
        }
        string _colorPath = "a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Sets the default color of the text.
        /// This sets the Fill to a SolidFill with the specified color.
        /// <remark>
        /// Use the Fill property for more options
        /// </remark>
        /// </summary>
        public override Color Color
        {
            get
            {
                return _textRun.Fill.Color;
            }
            set
            {
                _textRun.Fill.Style = eFillStyle.SolidFill;
                _textRun.Fill.SolidFill.Color.SetRgbColor(value);
            }
        }
        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run
        /// </summary>
        public override double Kerning
        {
            get
            {
                return _textRun.Kerning;
            }
            set
            {
                _textRun.Kerning = value;
            }
        }
        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public override  eTextCapsType Capitalization
        {
            get
            {
                return _textRun.Capitalization;
            }
            set
            {
                _textRun.Capitalization = value;
            }
        }

        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public override double Baseline
        {
            get
            {
                return _textRun.Baseline;
            }
            set
            {
                _textRun.Baseline = value;
            }
        }
        internal override bool IsEmpty => _textRun.IsEmpty;
        internal override XmlElement PathElement => (XmlElement)_textRun.TopNode;

    }
    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public abstract class ExcelTextFont 
    {
        /// <summary>
        /// The latin typeface name
        /// </summary>
        public abstract string LatinFont
        {
            get;
            set;
        }
        /// <summary>
        /// The East Asian typeface name
        /// </summary>
        public abstract string EastAsianFont
        {
            get;
            set;
        }
        /// <summary>
        /// The complex font typeface name
        /// </summary>
        public abstract string ComplexFont
        {
            get;
            set;
        }

        /// <summary>
        /// If the font is bold
        /// </summary>
        public abstract bool Bold
        {
            get;
            set;
        }
        /// <summary>
        /// The fonts underline style
        /// </summary>
        public abstract eUnderLineType UnderLine
        {
            get;
            set;
        }

        /// <summary>
        /// The fonts underline color
        /// </summary>
        public abstract Color UnderLineColor
        {
            get;
            set;
        }
        /// <summary>
        /// If the font is italic
        /// </summary>
        public abstract bool Italic
        {
            get;
            set;
        }
        /// <summary>
        /// Font strike out type
        /// </summary>
        public abstract eStrikeType Strike
        {
            get;
            set;
        }
        /// <summary>
        /// Font size
        /// </summary>
        public abstract float Size
        {
            get;
            set;
        }
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public abstract ExcelDrawingFill Fill
        {
            get;
        }
        /// <summary>
        /// Sets the default color of the text.
        /// This sets the Fill to a SolidFill with the specified color.
        /// <remark>
        /// Use the Fill property for more options
        /// </remark>
        /// </summary>
        public abstract Color Color
        {
            get;
            set;
        }
        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run
        /// </summary>
        public abstract double Kerning
        {
            get;
            set;
        }
        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public abstract eTextCapsType Capitalization
        {
            get;
            set;
        }

        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public abstract double Baseline
        {
            get;
            set;
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
            Size = size;
            if (bold) Bold = bold;
            if (italic) Italic = italic;
            if (underline) UnderLine = eUnderLineType.Single;
            if (strikeout) Strike = eStrikeType.Single;
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

        internal abstract bool IsEmpty
        {
            get;
        }
        internal abstract XmlElement PathElement
        {
            get;
        }        
    }
}
