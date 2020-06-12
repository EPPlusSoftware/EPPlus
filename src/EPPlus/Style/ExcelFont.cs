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
            Index = (index == int.MinValue ? 0 : index);
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
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Name, value, _positionID, _address));
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
                //_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.UnderlineType, value, _positionID, _address));
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
        /// Set the font from a Font object
        /// </summary>
        /// <param name="Font"></param>
        public void SetFromFont(Font Font)
        {
            Name = Font.Name;
            //Family=fnt.FontFamily.;
            Size = (int)Font.Size;
            Strike = Font.Strikeout;
            Bold = Font.Bold;
            UnderLine = Font.Underline;
            Italic = Font.Italic;
        }

        internal override string Id
        {
            get 
            {
                return Name + Size.ToString() + Family.ToString() + Scheme.ToString() + Bold.ToString()[0] + Italic.ToString()[0] + Strike.ToString()[0] + UnderLine.ToString()[0] + VerticalAlign;
            }
        }
    }
}
