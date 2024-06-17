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
using System.Xml;
using OfficeOpenXml.Drawing;
using System.Drawing;
using OfficeOpenXml.Drawing.Interfaces;

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
        }
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
        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return GetXmlNodeBool(_boldPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_boldPath, value ? "1" : "0");
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
                return GetXmlNodeString(_underLinePath).TranslateUnderline();
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_underLinePath, value.TranslateUnderlineText());
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
        string _italicPath = "@i";
        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return GetXmlNodeBool(_italicPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_italicPath, value ? "1" : "0");
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
                return GetXmlNodeString(_strikePath).TranslateStrikeType();
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_strikePath, value.TranslateStrikeTypeText());
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
                return GetXmlNodeInt(_sizePath) / 100;
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_sizePath, ((int)(value * 100)).ToString());
            }
        }
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
        string _colorPath = "a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Sets the default color of the text.
        /// This sets the Fill to a SolidFill with the specified color.
        /// <remark>
        /// Use the Fill property for more options
        /// </remark>
        /// </summary>
        [Obsolete("Use the Fill property for more options")]
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
        string _kernPath = "@kern";
        /// <summary>
        /// Specifies the minimum font size at which character kerning occurs for this text run
        /// </summary>
        public double Kerning
        {
            get
            {
                return GetXmlNodeFontSize(_kernPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeFontSize(_kernPath, value, "Kerning");
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
            Size = size;
            if (bold) Bold = bold;
            if (italic) Italic = italic;
            if (underline) UnderLine = eUnderLineType.Single;
            if (strikeout) Strike = eStrikeType.Single;            
        }
    }
}
