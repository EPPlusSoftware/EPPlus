using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingColorScaleValue
    {
        int _priority;

        ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectType? type,
            Color color,
            double value,
            string formula,
            int priority)
        {
            Type = (eExcelConditionalFormattingValueObjectType)type;
            _colorSettings = new ExcelDxfColor(null, eStyleClass.Fill, SetColor);
            Color = color;
            Value = value;
            Formula = formula;
            _priority = priority;
        }

        internal ExcelConditionalFormattingColorScaleValue(
        eExcelConditionalFormattingValueObjectType? type,
        Color color,
        int priority) 
        : this(type, color, double.NaN, null, priority)
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
        }

        internal void SetColor(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            if (styleProperty == eStyleProperty.Color)
            {
               _color = (Color)value;
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
                    _value = value;
                }
            }
        }

        string _formula;

        /// <summary>
        /// The Formula of the Object Value (uses the same attribute as the Value)
        /// </summary>
        public string Formula
        {
            get
            {
                // Return empty if the Object Value type is not Formula
                if (Type != eExcelConditionalFormattingValueObjectType.Formula)
                {
                    return string.Empty;
                }

                // Excel stores the formula in the @val attribute
                return _formula;
            }
            set
            {
                // Only store the formula if the Object Value type is Formula
                if (Type == eExcelConditionalFormattingValueObjectType.Formula)
                {
                   _formula = value;
                }
            }
        }
    }
}
