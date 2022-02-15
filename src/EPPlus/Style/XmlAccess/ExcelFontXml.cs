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
#if NETFULL
using System.Drawing;
#endif
using System.Globalization;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for fonts
    /// </summary>
    public sealed class ExcelFontXml : StyleXmlHelper
    {
        internal ExcelFontXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {
            _name = "";
            _size = 0;
            _family = int.MinValue;
            _scheme = "";
            _color = _color = new ExcelColorXml(NameSpaceManager);
            _bold = false;
            _italic = false;
            _strike = false;
            _underlineType = ExcelUnderLineType.None ;
            _verticalAlign = "";
            _charset = null;
        }
        internal ExcelFontXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _name = GetXmlNodeString(namePath);
            _size = (float)GetXmlNodeDecimal(sizePath);
            _family = GetXmlNodeIntNull(familyPath)??int.MinValue;
            _scheme = GetXmlNodeString(schemePath);
            _color = new ExcelColorXml(nsm, topNode.SelectSingleNode(_colorPath, nsm));
            _bold = GetBoolValue(topNode, boldPath);
            _italic = GetBoolValue(topNode, italicPath);
            _strike = GetBoolValue(topNode, strikePath);
            _verticalAlign = GetXmlNodeString(verticalAlignPath);
            _charset = GetXmlNodeIntNull(_charsetPath);
            if (topNode.SelectSingleNode(underLinedPath, NameSpaceManager) != null)
            {
                string ut = GetXmlNodeString(underLinedPath + "/@val");
                if (ut == "")
                {
                    _underlineType = ExcelUnderLineType.Single;
                }
                else
                {
                    _underlineType = (ExcelUnderLineType)Enum.Parse(typeof(ExcelUnderLineType), ut, true);
                }
            }
            else
            {
                _underlineType = ExcelUnderLineType.None;
            }
        }

        internal override string Id
        {
            get
            {
                return Name + "|" + Size + "|" + Family + "|" + Color.Id + "|" + Scheme + "|" + Bold.ToString() + "|" + Italic.ToString() + "|" + Strike.ToString() + "|" + VerticalAlign + "|" + UnderLineType.ToString() + "|" + (Charset.HasValue ? Charset.ToString() : "");
            }
        }
        const string namePath = "d:name/@val";
        string _name;
        /// <summary>
        /// The name of the font
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                Scheme = "";        //Reset schema to avoid corrupt file if unsupported font is selected.
                _name = value;
            }
        }
        const string sizePath = "d:sz/@val";
        float _size;
        /// <summary>
        /// Font size
        /// </summary>
        public float Size
        {
            get
            {
                return _size;
            }
            set
            {
                _size = value;
            }
        }
        const string familyPath = "d:family/@val";
        int _family;
        /// <summary>
        /// Font family
        /// </summary>
        public int Family
        {
            get
            {
                return (_family == int.MinValue ? 0 : _family);
            }
            set
            {
                _family=value;
            }
        }
        ExcelColorXml _color = null;
        const string _colorPath = "d:color";
        /// <summary>
        /// Text color
        /// </summary>
        public ExcelColorXml Color
        {
            get
            {
                return _color;
            }
            internal set 
            {
                _color=value;
            }
        }
        const string schemePath = "d:scheme/@val";
        string _scheme="";
        /// <summary>
        /// Font Scheme
        /// </summary>
        public string Scheme
        {
            get
            {
                return _scheme;
            }
            internal set
            {
                _scheme=value;
            }
        }
        const string boldPath = "d:b";
        bool _bold;
        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return _bold;
            }
            set
            {
                _bold=value;
            }
        }
        const string italicPath = "d:i";
        bool _italic;
        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return _italic;
            }
            set
            {
                _italic=value;
            }
        }
        const string strikePath = "d:strike";
        bool _strike;
        /// <summary>
        /// If the font is striked out
        /// </summary>
        public bool Strike
        {
            get
            {
                return _strike;
            }
            set
            {
                _strike=value;
            }
        }
        const string underLinedPath = "d:u";
        /// <summary>
        /// If the font is underlined.
        /// When set to true a the text is underlined with a single line
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return UnderLineType!=ExcelUnderLineType.None;
            }
            set
            {
                _underlineType=value ? ExcelUnderLineType.Single : ExcelUnderLineType.None;
            }
        }
        ExcelUnderLineType _underlineType;
        /// <summary>
        /// If the font is underlined
        /// </summary>
        public ExcelUnderLineType UnderLineType
        {
            get
            {
                return _underlineType;
            }
            set
            {
                _underlineType = value;
            }
        }
        const string verticalAlignPath = "d:vertAlign/@val";
        string _verticalAlign;
        /// <summary>
        /// Vertical aligned
        /// </summary>
        public string VerticalAlign
        {
            get
            {
                return _verticalAlign;
            }
            set
            {
                _verticalAlign=value;
            }
        }
        const string _charsetPath = "d:charset/@val";
        int? _charset=null;
        /// <summary>
        /// The character set for the font
        /// </summary>
        /// <remarks>
        /// The following values can be used for this property.
        /// <list type="table">
        /// <listheader>Value</listheader><listheader>Description</listheader>
        /// <item>null</item><item>Not specified</item>
        /// <item>0x00</item><item>The ANSI character set. (IANA name iso-8859-1)</item>
        /// <item>0x01</item><item>The default character set.</item>
        /// <item>0x02</item><item>The Symbol character set. This value specifies that the characters in the Unicode private use area(U+FF00 to U+FFFF) of the font should be used to display characters in the range U+0000 to U+00FF.</item>       
        ///<item>0x4D</item><item>A Macintosh(Standard Roman) character set. (IANA name macintosh)</item>
        ///<item>0x80</item><item>The JIS character set. (IANA name shift_jis)</item>
        ///<item>0x81</item><item>The Hangul character set. (IANA name ks_c_5601-1987)</item>
        ///<item>0x82</item><item>A Johab character set. (IANA name KS C-5601-1992)</item>
        ///<item>0x86</item><item>The GB-2312 character set. (IANA name GBK)</item>
        ///<item>0x88</item><item>The Chinese Big Five character set. (IANA name Big5)</item>
        ///<item>0xA1</item><item>A Greek character set. (IANA name windows-1253)</item>
        ///<item>0xA2</item><item>A Turkish character set. (IANA name iso-8859-9)</item>
        ///<item>0xA3</item><item>A Vietnamese character set. (IANA name windows-1258)</item>
        ///<item>0xB1</item><item>A Hebrew character set. (IANA name windows-1255)</item>
        ///<item>0xB2</item><item>An Arabic character set. (IANA name windows-1256)</item>
        ///<item>0xBA</item><item>A Baltic character set. (IANA name windows-1257)</item>
        ///<item>0xCC</item><item>A Russian character set. (IANA name windows-1251)</item>
        ///<item>0xDE</item><item>A Thai character set. (IANA name windows-874)</item>
        ///<item>0xEE</item><item>An Eastern European character set. (IANA name windows-1250)</item>
        ///<item>0xFF</item><item>An OEM character set not defined by ISO/IEC 29500.</item>
        ///<item>Any other value</item><item>Application-defined, can be ignored</item>
        /// </list>
        /// </remarks>
        public int? Charset
        {
            get
            {
                return _charset;
            }
            set
            {
                _charset = value;
            }
        }

#if NETFULL
        [Obsolete("This method is deprecated and is removed .NET standard/core. Please use overloads not referencing System.Drawing.Font")]
        public void SetFromFont(Font Font)
        {
            SetFromFont(Font.Name, Font.Size, Font.Bold, Font.Italic, Font.Underline, Font.Strikeout);
        }
#endif

        /// <summary>
        /// Set the font properties
        /// </summary>
        /// <param name="name">Font family name</param>
        /// <param name="size">Font size</param>
        /// <param name="bold"></param>
        /// <param name="italic"></param>
        /// <param name="underline"></param>
        /// <param name="strikeout"></param>
        public void SetFromFont(string name, float size, bool bold = false, bool italic = false, bool underline = false, bool strikeout = false)
        {
            Name=name;
            //Family=fnt.FontFamily.;
            Size= size;
            Strike= strikeout;
            Bold = bold;
            UnderLine= underline;
            Italic= italic;            
        }
        /// <summary>
        /// Gets the height of the font in 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        internal static float GetFontHeight(string name, float size)
        {
            name = name.StartsWith("@") ? name.Substring(1) : name;
            return Convert.ToSingle(ExcelWorkbook.GetHeightPixels(name, size));
            if(FontSize.FontHeights.ContainsKey(name)==false)
            {
                return GetHeightByName("Calibri", size);
            }
        }
        internal ExcelFontXml Copy()
        {
            ExcelFontXml newFont = new ExcelFontXml(NameSpaceManager);
            newFont.Name = _name;
            newFont.Size = _size;
            newFont.Family = _family;
            newFont.Scheme = _scheme;
            newFont.Bold = _bold;
            newFont.Italic = _italic;
            newFont.UnderLineType = _underlineType;
            newFont.Strike = _strike;
            newFont.VerticalAlign = _verticalAlign;
            newFont.Color = Color.Copy();
            newFont.Charset = _charset;
            return newFont;
        }
        internal override XmlNode CreateXmlNode(XmlNode topElement)
        {
            TopNode = topElement;
            if (_bold) CreateNode(boldPath); else DeleteAllNode(boldPath);
            if (_italic) CreateNode(italicPath); else DeleteAllNode(italicPath);
            if (_strike) CreateNode(strikePath); else DeleteAllNode(strikePath);
            
            if (_underlineType == ExcelUnderLineType.None)
            {
                DeleteAllNode(underLinedPath);
            }
            else if(_underlineType==ExcelUnderLineType.Single)
            {
                CreateNode(underLinedPath);
            }
            else
            {
                var v=_underlineType.ToString();
                SetXmlNodeString(underLinedPath + "/@val", v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1));
            }

            if (_verticalAlign!="") SetXmlNodeString(verticalAlignPath, _verticalAlign.ToString());
            if(_size>0) SetXmlNodeString(sizePath, _size.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (_color.Exists)
            {
                CreateNode(_colorPath);
                TopNode.AppendChild(_color.CreateXmlNode(TopNode.SelectSingleNode(_colorPath, NameSpaceManager)));
            }

            if (!string.IsNullOrEmpty(_name)) SetXmlNodeString(namePath, _name);
            if(_family>int.MinValue) SetXmlNodeString(familyPath, _family.ToString());
            SetXmlNodeInt(_charsetPath, Charset);
            if (_scheme != "") SetXmlNodeString(schemePath, _scheme.ToString());

            return TopNode;
        }
    }
}
