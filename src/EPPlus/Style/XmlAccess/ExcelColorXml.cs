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
using System.Globalization;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for color
    /// </summary>
    public sealed class ExcelColorXml : StyleXmlHelper
    {
        internal ExcelColorXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {
            _auto = false;
            _theme = null;
            _tint = 0;
            _rgb = "";
            _indexed = int.MinValue;
        }
        internal ExcelColorXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            if(topNode==null)
            {
                Exists=false;
            }
            else
            {
                Exists = true;
                _auto = GetXmlNodeBool("@auto");
                var v=GetXmlNodeIntNull("@theme");
                if(v.HasValue && v>=0 && v<=11)
                {
                    _theme = (eThemeSchemeColor)v;
                }
                _tint = GetXmlNodeDecimalNull("@tint")??decimal.MinValue;
                _rgb = GetXmlNodeString("@rgb");
                _indexed = GetXmlNodeIntNull("@indexed") ?? int.MinValue;
            }
        }
        
        internal override string Id
        {
            get
            {
                return _auto.ToString() + "|" + _theme?.ToString() + "|" + _tint + "|" + _rgb + "|" + _indexed;
            }
        }
        bool _auto;
        /// <summary>
        /// Set the color to automatic
        /// </summary>
        public bool Auto
        {
            get
            {
                return _auto;
            }
            set
            {
                Clear();
                _auto = value;
                Exists = true;
            }
        }
        eThemeSchemeColor? _theme;
        /// <summary>
        /// Theme color value
        /// </summary>
        public eThemeSchemeColor? Theme
        {
            get
            {
                return _theme;
            }
            set
            {
                Clear();
                _theme = value;
                Exists = true;
            }
        }
        decimal _tint;
        /// <summary>
        /// The Tint value for the color
        /// </summary>
        public decimal Tint
        {
            get
            {
                if (_tint == decimal.MinValue)
                {
                    return 0;
                }
                else
                {
                    return _tint;
                }
            }
            set
            {
                _tint = value;
                Exists = true;
            }
        }
        string _rgb;
        /// <summary>
        /// The RGB value
        /// </summary>
        public string Rgb
        {
            get
            {
                return _rgb;
            }
            set
            {
                _rgb = value;
                Exists=true;
                _indexed = int.MinValue;
                _auto = false;
            }
        }
        int _indexed;
        /// <summary>
        /// Indexed color value.
        /// Returns int.MinValue if indexed colors are not used.
        /// </summary>
        public int Indexed
        {
            get
            {
                return _indexed;
            }
            set
            {
                if (value < 0 || value > 65)
                {
                    throw (new ArgumentOutOfRangeException("Index out of range"));
                }
                Clear();
                _indexed = value;
                Exists = true;
            }
        }
        internal void Clear()
        {
            _theme = null;
            _tint = decimal.MinValue;
            _indexed = int.MinValue;
            _rgb = "";
            _auto = false;
        }
        /// <summary>
        /// Sets the color
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(System.Drawing.Color color)
        {
            Clear();
            _rgb = color.ToArgb().ToString("X");
        }
        /// <summary>
        /// Sets a theme color
        /// </summary>
        /// <param name="themeColorType">The theme color</param>
        public void SetColor(eThemeSchemeColor themeColorType)
        {
            Clear();
            _theme = themeColorType;
        }
        /// <summary>
        /// Sets an indexed color
        /// </summary>
        /// <param name="indexedColor">The indexed color</param>
        public void SetColor(ExcelIndexedColor indexedColor)
        {
            Clear();
            _indexed = (int)indexedColor;
        }

        internal ExcelColorXml Copy()
        {
            return new ExcelColorXml(NameSpaceManager) {_indexed=_indexed, _tint=_tint, _rgb=_rgb, _theme=_theme, _auto=_auto, Exists=Exists };
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            if(_rgb!="")
            {
                SetXmlNodeString("@rgb", _rgb);
            }
            else if (_indexed >= 0)
            {
                SetXmlNodeString("@indexed", _indexed.ToString());
            }
            else if (_auto)
            {
                SetXmlNodeBool("@auto", _auto);
            }
            else
            {
                SetXmlNodeString("@theme", ((int)_theme).ToString(CultureInfo.InvariantCulture));
            }
            if (_tint != decimal.MinValue)
            {
                SetXmlNodeString("@tint", _tint.ToString(CultureInfo.InvariantCulture));
            }
            return TopNode;
        }
        /// <summary>
        /// True if the record exists in the underlaying xml
        /// </summary>
        internal bool Exists { get; private set; }
    }
}
