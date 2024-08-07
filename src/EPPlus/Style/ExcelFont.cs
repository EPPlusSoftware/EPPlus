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
using System.Text;
using System.Drawing;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Cell style Font
    /// </summary>
    public sealed class ExcelFont : StyleBase
    {
        internal ExcelFont(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)

        {
            Index = index;
        }
        /// <summary>
        /// The name of the font
        /// </summary>
        public string Name
        {
            get
            {
                return _styles.Fonts[Index].Name;
            }
            set
            {
                CheckNormalStyleChange();
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Name, value, _positionID, _address));
            }
        }

        private void CheckNormalStyleChange()
        {
            var nsIx = _styles.GetNormalStyleIndex();
            if(nsIx>=0)
            {
                if(_styles.NamedStyles[nsIx].Style.Font.Index==Index)
                {
                    _styles._wb.ClearDefaultHeightsAndWidths();
                }
            }
        }

        /// <summary>
        /// The Size of the font
        /// </summary>
        public float Size
        {
            get
            {
                return _styles.Fonts[Index].Size;
            }
            set
            {
                CheckNormalStyleChange();
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Size, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font family
        /// </summary>
        public int Family
        {
            get
            {
                return _styles.Fonts[Index].Family;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Family, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Cell color
        /// </summary>
        public ExcelColor Color
        {
            get
            {
                return new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.Font, this);
            }
        }
        /// <summary>
        /// Scheme
        /// </summary>
        public string Scheme
        {
            get
            {
                return _styles.Fonts[Index].Scheme;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Scheme, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return _styles.Fonts[Index].Bold;
            }
            set
            {
                CheckNormalStyleChange();
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Bold, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return _styles.Fonts[Index].Italic;
            }
            set
            {
                CheckNormalStyleChange();
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Italic, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-Strikeout
        /// </summary>
        public bool Strike
        {
            get
            {
                return _styles.Fonts[Index].Strike;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Strike, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-Underline
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return _styles.Fonts[Index].UnderLine;
            }
            set
            {
                if (value)
                {
                    UnderLineType = ExcelUnderLineType.Single;
                }
                else
                {
                    UnderLineType = ExcelUnderLineType.None;
                }
                //_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.UnderlineType, value, _positionID, _addresses));
            }
        }
        /// <summary>
        /// The underline style
        /// </summary>
        public ExcelUnderLineType UnderLineType
        {
            get
            {
                return _styles.Fonts[Index].UnderLineType;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.UnderlineType, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Font-Vertical Align
        /// </summary>
        public ExcelVerticalAlignmentFont VerticalAlign
        {
            get
            {
                if (_styles.Fonts[Index].VerticalAlign == "")
                {
                    return ExcelVerticalAlignmentFont.None;
                }
                else
                {
                    return (ExcelVerticalAlignmentFont)Enum.Parse(typeof(ExcelVerticalAlignmentFont), _styles.Fonts[Index].VerticalAlign, true);
                }
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.VerticalAlign, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The character set for the font
        /// </summary>
        /// <remarks>
        /// The following values can be used for this property
        /// <list type="table">f
        /// <listheader>Value</listheader><listheader>Description</listheader>
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
                return _styles.Fonts[Index].Charset;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Charset, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Set the font from a Font object
        /// </summary>
        /// <param name="name">Font family name</param>
        /// <param name="bold"></param>
        /// <param name="size">Font size</param>
        /// <param name="italic"></param>
        /// <param name="underline"></param>
        /// <param name="strikeout"></param>
        public void SetFromFont(string name, float size, bool bold = false, bool italic = false, bool underline = false, bool strikeout = false)
        {
            Name = name;
            Size = size;
            Strike = strikeout;
            Bold = bold;
            UnderLine = underline;
            Italic = italic;
        }

        internal override string Id
        {
            get 
            {
                return Name + Size.ToString() + Family.ToString() + Scheme.ToString() + Bold.ToString()[0] + Italic.ToString()[0] + Strike.ToString()[0] + UnderLine.ToString()[0] + VerticalAlign + Charset.ToString();
            }
        }
    }
}
