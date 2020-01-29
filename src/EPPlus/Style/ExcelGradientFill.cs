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
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// The background fill of a cell
    /// </summary>
    public class ExcelGradientFill : StyleBase
    {
        internal ExcelGradientFill(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)

        {
            Index = index;
        }
        /// <summary>
        /// Angle of the linear gradient
        /// </summary>
        public double Degree
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Degree;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientDegree, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Linear or Path gradient
        /// </summary>
        public ExcelFillGradientType Type
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Type;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientType, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The top position of the inner rectangle (color 1) in percentage format (from the top to the bottom). 
        /// Spans from 0 to 1
        /// </summary>
        public double Top
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Top;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientTop, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The bottom position of the inner rectangle (color 1) in percentage format (from the top to the bottom). 
        /// Spans from 0 to 1
        /// </summary>
        public double Bottom
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Bottom;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientBottom, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The left position of the inner rectangle (color 1) in percentage format (from the left to the right). 
        /// Spans from 0 to 1
        /// </summary>
        public double Left
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Left;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientLeft, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The right position of the inner rectangle (color 1) in percentage format (from the left to the right). 
        /// Spans from 0 to 1
        /// </summary>
        public double Right
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Right;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientRight, value, _positionID, _address));
            }
        }
        ExcelColor _gradientColor1 = null;
        /// <summary>
        /// Gradient Color 1
        /// </summary>
        public ExcelColor Color1
        {
            get
            {
                if (_gradientColor1 == null)
                {
                    _gradientColor1 = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillGradientColor1, this);
                }
                return _gradientColor1;

            }
        }
        ExcelColor _gradientColor2 = null;
        /// <summary>
        /// Gradient Color 2
        /// </summary>
        public ExcelColor Color2
        {
            get
            {
                if (_gradientColor2 == null)
                {
                    _gradientColor2 = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillGradientColor2, this);
                }
                return _gradientColor2;

            }
        }
        internal override string Id
        {
            get { return Degree.ToString() + Type + Color1.Id + Color2.Id + Top.ToString() + Bottom.ToString() + Left.ToString() + Right.ToString(); }
        }
    }
}
