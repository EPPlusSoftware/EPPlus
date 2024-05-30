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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Drawing;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelConditionalFormattingColorScaleValue
    {
        int _priority;

        ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectType? type,
            Color color,
            double value,
            string formula,
            int priority,
            ExcelStyles styles)
        {
            Type = (eExcelConditionalFormattingValueObjectType)type;
            _colorSettings = new ExcelDxfColor(styles, eStyleClass.Fill, SetColor);
            Color = color;
            Value = value;
            if(Type != eExcelConditionalFormattingValueObjectType.Percentile)
            {
                Formula = formula;
            }
            _priority = priority;
        }

        internal ExcelConditionalFormattingColorScaleValue(
        eExcelConditionalFormattingValueObjectType? type,
        Color color,
        int priority, ExcelStyles styles) 
        : this(type, color, double.NaN, null, priority, styles)
        {
        }

        /// <summary>
        /// The value type
        /// </summary>
        public eExcelConditionalFormattingValueObjectType Type{ get; set; }

        ExcelDxfColor _colorSettings;
        Color _color;

        /// <summary>
        /// Used to set color or theme color, index, auto and tint
        /// </summary>
        public ExcelDxfColor ColorSettings
        {
            get 
            { 
                return _colorSettings;
            }
            internal set 
            { 
                _colorSettings = value;
            }
        }

        internal void SetColor(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            if (styleProperty == eStyleProperty.Color)
            {
                if(value != null)
                {
                    _color = (Color)value;
                }
                else
                {
                    _color = Color.Empty;
                }
            }
        }

        /// <summary>
        /// The color to be used
        /// </summary>
        public Color Color
        {
            get
            {
                return _color;
            }
            set
            {
                _color = value;
                ColorSettings.SetColor(value);
            }
        }

        Double _value = double.NaN;

        /// <summary>
        /// The value of the conditional formatting
        /// </summary>
        public Double Value
        {
            get
            {
                return _value;
            }
            set
            {
                // Only some types use the @val attribute
                if ((Type == eExcelConditionalFormattingValueObjectType.Num)
                    || (Type == eExcelConditionalFormattingValueObjectType.Percent)
                    || (Type == eExcelConditionalFormattingValueObjectType.Percentile))
                {
                    _formula = null;
                    _value = value;
                }
            }
        }

        string _formula;

        /// <summary>
        /// <para> The Formula of the Object Value </para>
        /// Keep in mind that Addresses in this property should be Absolute not relative  
        /// <para> Yes: $A$1 </para> 
        /// <para> No: A1 </para>
        /// </summary>
        public string Formula
        {
            get
            {
                // Return empty if the Object Value type is not Formula
                if (Type == eExcelConditionalFormattingValueObjectType.Percentile)
                {
                    return string.Empty;
                }

                // Excel stores the formula in the @val attribute
                return _formula;
            }
            set
            {
                // Only store the formula if the Object Value type is Formula
                if (Type != eExcelConditionalFormattingValueObjectType.Percentile)
                {
                    _value = double.NaN;
                    _formula = value;
                }
                else
                {
                    throw new InvalidOperationException("Cannot store formula in a percentile type.");
                }
            }
        }
    }
}
