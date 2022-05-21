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
using OfficeOpenXml.Style.XmlAccess;
using System.Drawing;
using OfficeOpenXml.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Color for cellstyling
    /// </summary>
    public sealed class ExcelColor :  StyleBase, IColor
    {
        internal static string[] indexedColors =
        {
                "#FF000000", // 0
                "#FFFFFFFF",
                "#FFFF0000",
                "#FF00FF00",
                "#FF0000FF",
                "#FFFFFF00",
                "#FFFF00FF",
                "#FF00FFFF",
                "#FF000000", // 8
                "#FFFFFFFF",
                "#FFFF0000",
                "#FF00FF00",
                "#FF0000FF",
                "#FFFFFF00",
                "#FFFF00FF",
                "#FF00FFFF",
                "#FF800000",
                "#FF008000",
                "#FF000080",
                "#FF808000",
                "#FF800080",
                "#FF008080",
                "#FFC0C0C0",
                "#FF808080",
                "#FF9999FF",
                "#FF993366",
                "#FFFFFFCC",
                "#FFCCFFFF",
                "#FF660066",
                "#FFFF8080",
                "#FF0066CC",
                "#FFCCCCFF",
                "#FF000080",
                "#FFFF00FF",
                "#FFFFFF00",
                "#FF00FFFF",
                "#FF800080",
                "#FF800000",
                "#FF008080",
                "#FF0000FF",
                "#FF00CCFF",
                "#FFCCFFFF",
                "#FFCCFFCC",
                "#FFFFFF99",
                "#FF99CCFF",
                "#FFFF99CC",
                "#FFCC99FF",
                "#FFFFCC99",
                "#FF3366FF",
                "#FF33CCCC",
                "#FF99CC00",
                "#FFFFCC00",
                "#FFFF9900",
                "#FFFF6600",
                "#FF666699",
                "#FF969696",
                "#FF003366",
                "#FF339966",
                "#FF003300",
                "#FF333300",
                "#FF993300",
                "#FF993366",
                "#FF333399",
                "#FF333333", // 63
            };

        internal static Color GetIndexedColor(int index)
        {
            if(index >= 0 && index < indexedColors.Length)
            {
                var s = indexedColors[index];
                var a = int.Parse(s.Substring(1, 2), NumberStyles.HexNumber);
                var r = int.Parse(s.Substring(3, 2), NumberStyles.HexNumber);
                var g = int.Parse(s.Substring(5, 2), NumberStyles.HexNumber);
                var b = int.Parse(s.Substring(7, 2), NumberStyles.HexNumber);

                return Color.FromArgb(a, r, g, b);
            }
            return Color.Empty;
        }

        eStyleClass _cls;
        StyleBase _parent;
        internal ExcelColor(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent) : 
            base(styles, ChangedEvent, worksheetID, address)
        {
            _parent = parent;
            _cls = cls;
            Index = int.MinValue;
        }
        /// <summary>
        /// The theme color
        /// </summary>
        public eThemeSchemeColor? Theme
        {
            get
            {
                if (_parent.Index < 0) return null;
                return GetSource().Theme;
            }
            internal set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Theme, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The tint value
        /// </summary>
        public decimal Tint
        {
            get
            {
                if (_parent.Index < 0) return 0;
                return GetSource().Tint;
            }
            set
            {
                if (value > 1 || value < -1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between -1 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Tint, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The RGB value
        /// </summary>
        public string Rgb
        {
            get
            {
                if (_parent.Index < 0) return null;
                return GetSource().Rgb;
            }
            internal set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Color, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The indexed color number.
        /// A negative value means not set.
        /// </summary>
        public int Indexed
        {
            get
            {
                if (_parent.Index < 0) return -1;
                return GetSource().Indexed;
            }
            set
            {
                if(value<0)
                {
                    throw (new ArgumentOutOfRangeException("Indexed", "Cannot be negative"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.IndexedColor, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Auto color
        /// </summary>
        public bool Auto
        {
            get
            {
                if (_parent.Index < 0) return false;

                return GetSource().Auto;
            }
            private set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.AutoColor, value, _positionID, _address));
            }
        }

        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(Color color)
        {
            Rgb = color.ToArgb().ToString("X");       
        }
        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(eThemeSchemeColor color)
        {
            Theme=color;
        }
        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(ExcelIndexedColor color)
        {
            Indexed = (int)color;
        }
        /// <summary>
        /// Set the color to automatic
        /// </summary>
        public void SetAuto()
        {
            Auto = true;
        }
        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="alpha">Alpha component value</param>
        /// <param name="red">Red component value</param>
        /// <param name="green">Green component value</param>
        /// <param name="blue">Blue component value</param>
        public void SetColor(int alpha, int red, int green, int blue)
        {
            if(alpha < 0 || red < 0 || green < 0 ||blue < 0 ||
               alpha > 255 || red > 255 || green > 255 || blue > 255)
            {
                throw (new ArgumentException("Argument range must be from 0 to 255"));
            }
            Rgb = alpha.ToString("X2") + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
        }
        internal override string Id
        {
            get 
            {
                return Theme.ToString() + Tint + Rgb + Indexed;
            }
        }

        private ExcelColorXml GetSource()
        {
            Index = _parent.Index < 0 ? 0 : _parent.Index;
            switch (_cls)
            {
                case eStyleClass.FillBackgroundColor:
                    return _styles.Fills[Index].BackgroundColor;
                case eStyleClass.FillPatternColor:
                    return _styles.Fills[Index].PatternColor;
                case eStyleClass.Font:
                    return _styles.Fonts[Index].Color;
                case eStyleClass.BorderLeft:
                    return _styles.Borders[Index].Left.Color;
                case eStyleClass.BorderTop:
                    return _styles.Borders[Index].Top.Color;
                case eStyleClass.BorderRight:
                    return _styles.Borders[Index].Right.Color;
                case eStyleClass.BorderBottom:
                    return _styles.Borders[Index].Bottom.Color;
                case eStyleClass.BorderDiagonal:
                    return _styles.Borders[Index].Diagonal.Color;
                case eStyleClass.FillGradientColor1:
                    return ((ExcelGradientFillXml)(_styles.Fills[Index])).GradientColor1;
                case eStyleClass.FillGradientColor2:
                    return ((ExcelGradientFillXml)(_styles.Fills[Index])).GradientColor2;
                default:
                    throw(new Exception("Invalid style-class for Color"));
            }
        }
        internal override void SetIndex(int index)
        {
            _parent.Index = index;
        }
        /// <summary>
        /// Return the RGB hex string for the Indexed or Tint property
        /// </summary>
        /// <returns>The RGB color starting with a #FF (alpha)</returns>
        public string LookupColor()
        {
            return LookupColor(this);
        }
        /// <summary>
        /// Return the RGB value as a string for the color object that uses the Indexed or Tint property
        /// </summary>
        /// <param name="theColor">The color object</param>
        /// <returns>The RGB color starting with a #FF (alpha)</returns>
        public string LookupColor(ExcelColor theColor)
        {
            if (theColor.Indexed >= 0 && theColor.Indexed < indexedColors.Length)
            {
                return indexedColors[theColor.Indexed];
            }
            else if (theColor.Rgb != null && theColor.Rgb.Length > 0)
            {
                return "#" + theColor.Rgb;
            }
            else
            {
                var c = ((int)(Math.Round((theColor.Tint+1) * 128))).ToString("X");
                return "#FF" + c + c + c;
            }
        }
    }
}
