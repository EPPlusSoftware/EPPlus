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
using System.Globalization;
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// The background fill of a cell
    /// </summary>
    public class ExcelFill : StyleBase
    {
        internal ExcelFill(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)
        {
            Index = index;
        }
        /// <summary>
        /// The pattern for solid fills.
        /// </summary>
        public ExcelFillStyle PatternType
        {
            get
            {
                if (Index == int.MinValue)
                {
                    return ExcelFillStyle.None;
                }
                else
                {
                    return _styles.Fills[Index].PatternType;
                }
            }
            set
            {
                if (_gradient != null) _gradient = null;
                    _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Fill, eStyleProperty.PatternType, value, _positionID, _address));
            }
        }
        ExcelColor _patternColor = null;
        /// <summary>
        /// The color of the pattern
        /// </summary>
        public ExcelColor PatternColor
        {
            get
            {
                if (_patternColor == null)
                {
                    _patternColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillPatternColor, this);
                    if (_gradient != null) _gradient = null;
                }
                return _patternColor;
            }
        }
        ExcelColor _backgroundColor = null;
        /// <summary>
        /// The background color
        /// </summary>
        public ExcelColor BackgroundColor
        {
            get
            {
                if (_backgroundColor == null)
                {
                    _backgroundColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillBackgroundColor, this);
                    if (_gradient != null) _gradient = null;
                }
                return _backgroundColor;
                
            }
        }
        ExcelGradientFill _gradient=null;
        /// <summary>
        /// Access to properties for gradient fill.
        /// </summary>
        public ExcelGradientFill Gradient 
        {
            get
            {
                if (_gradient == null)
                {                    
                    _gradient = new ExcelGradientFill(_styles, _ChangedEvent, _positionID, _address, Index);
                    _backgroundColor = null;
                    _patternColor = null;
                }
                return _gradient;
            }
        }
        internal override string Id
        {
            get
            {
                if (_gradient == null)
                {
                    return PatternType + PatternColor.Id + BackgroundColor.Id;
                }
                else
                {
                    return _gradient.Id;
                }
            }
        }

        //public void SetBackgroundAndPatternColor(Color bgColor, Color fgColor)
        //{
        //    PatternColor.SetColor(bgColor);
        //    BackgroundColor.SetColor(fgColor);
        //}

        /// <summary>
        /// Set the background to a specific color and fillstyle
        /// </summary>
        /// <param name="color">the color</param>
        /// <param name="fillStyle">The fillstyle. Default Solid</param>
        public void SetBackground(Color color, ExcelFillStyle fillStyle=ExcelFillStyle.Solid)
        {
            PatternType = fillStyle;
            BackgroundColor.SetColor(color);
        }
        /// <summary>
        /// Set the background to a specific color and fillstyle
        /// </summary>
        /// <param name="color">The indexed color</param>
        /// <param name="fillStyle">The fillstyle. Default Solid</param>
        public void SetBackground(ExcelIndexedColor color, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            PatternType = fillStyle;
            BackgroundColor.SetColor(color);
        }
        /// <summary>
        /// Set the background to a specific color and fillstyle
        /// </summary>
        /// <param name="color">The theme color</param>
        /// <param name="fillStyle">The fillstyle. Default Solid</param>
        public void SetBackground(eThemeSchemeColor color, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            PatternType = fillStyle;
            BackgroundColor.SetColor(color);
        }
    }
}
    