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
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Globalization;
using System;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Direction of Databar
    /// </summary>
    public enum eDatabarDirection
    {
        /// <summary>
        /// Based on context
        /// </summary>
        Context = 0,
        /// <summary>
        /// Databar going from left to right
        /// </summary>
        LeftToRight = 1,
        /// <summary>
        /// Databar going RighToLeft
        /// </summary>
        RightToLeft = 2
    }

internal class ExcelConditionalFormattingDataBar : ExcelConditionalFormattingRule,
        IExcelConditionalFormattingDataBarGroup
    {
        internal ExcelConditionalFormattingDataBar(
         ExcelAddress address,
         int priority,
         ExcelWorksheet ws)
        : base(eExcelConditionalFormattingRuleType.DataBar, address, priority, ws)
        {
            HighValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Max, eExcelConditionalFormattingRuleType.DataBar);
            LowValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Min, eExcelConditionalFormattingRuleType.DataBar);

            Uid = NewId();

            InitalizeDxfColours();

            Style.Fill.Style = eDxfFillStyle.GradientFill;

            //Excel default blue?
            FillColor.Color = Color.FromArgb(int.Parse("FF638EC6", NumberStyles.HexNumber));

            NegativeFillColor.Color = Color.Red;
            NegativeBorderColor.Color = Color.Red;
        }

        private void InitalizeDxfColours()
        {
            FillColor = new ExcelDxfColor(null, eStyleClass.Fill, null);
            BorderColor = new ExcelDxfColor(null, eStyleClass.Border, ValueWasSet);
            NegativeFillColor = new ExcelDxfColor(null, eStyleClass.Fill, ValueWasSet);
            NegativeBorderColor = new ExcelDxfColor(null, eStyleClass.Border, ValueWasSet);
            AxisColor = new ExcelDxfColor(null, eStyleClass.Border, null);
        }

        internal void ValueWasSet(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            if(styleClass == eStyleClass.Border)
            {
                Border = true;
                if(NegativeBorderColor.HasValue)
                {
                    NegativeBarBorderColorSameAsPositive = false;
                }
            }

            if(styleClass == eStyleClass.Fill)
            {
                NegativeBarColorSameAsPositive = false;
            }
        }

        internal ExcelConditionalFormattingDataBar(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.DataBar, address, ws, xr)
        {
            xr.Read();
            var highType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>().Value;
            HighValue = new ExcelConditionalFormattingIconDataBarValue(highType, eExcelConditionalFormattingRuleType.DataBar);

            if(!string.IsNullOrEmpty(xr.GetAttribute("val")))
            {
                if(highType != eExcelConditionalFormattingValueObjectType.Formula)
                {
                    HighValue.Value = Double.Parse(xr.GetAttribute("val"));
                }
            }

            xr.Read();
            var lowType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>().Value;
            LowValue = new ExcelConditionalFormattingIconDataBarValue(lowType, eExcelConditionalFormattingRuleType.DataBar);

            if (!string.IsNullOrEmpty(xr.GetAttribute("val")))
            {
                if(lowType != eExcelConditionalFormattingValueObjectType.Formula)
                {
                    LowValue.Value = Double.Parse(xr.GetAttribute("val"));
                }
            }

            xr.Read();

            InitalizeDxfColours();

            ReadInCTColor(xr, "fillColor");

            //enter databar exit node ->(local) extLst -> ext -> id
            xr.Read();
            xr.Read();
            xr.Read();

            _uid = xr.ReadString();

            // /ext -> /extLst
            xr.Read();
            xr.Read();
            xr.Read();
        }

        /// <summary>
        /// For reading all Databar CT_Colors Recursively until we hit a non-color node.
        /// </summary>
        /// <param name="xr"></param>
        /// <param name="altName">To force the color to write to. Useful e.g. when loading the local databar node that denotes fill color is just named Color</param>
        /// <exception cref="Exception"></exception>
        internal void ReadInCTColor(XmlReader xr, string altName = null)
        {
            ExcelDxfColor col;
            string nodeName = altName != null ? altName : xr.LocalName;

            switch (nodeName)
            {
                case "fillColor":
                    col = FillColor;
                break;

                case "borderColor":
                    col = BorderColor;
                break;

                case "negativeFillColor":
                    col = NegativeFillColor;
                break;

                case "negativeBorderColor":
                    col = NegativeBorderColor;
                break;

                case "axisColor":
                    col = AxisColor;
                break;
                
                default: throw new Exception($"{xr.LocalName} is not a CT_Color node and cannot be read.");
            }


            if (!string.IsNullOrEmpty(xr.GetAttribute("auto")))
            {
                col.Auto = xr.GetAttribute("auto") == "1" ? true : false;
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("theme")))
            {
                col.Theme = (eThemeSchemeColor)int.Parse(xr.GetAttribute("theme"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("indexed")))
            {
                col.Index = int.Parse(xr.GetAttribute("indexed"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("rgb")))
            {
                col.Color = (ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb")));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("tint")))
            {
                col.Tint = double.Parse(xr.GetAttribute("tint"));
            }

            xr.Read();

            if(xr.LocalName.Contains("Color"))
            {
                ReadInCTColor(xr);
            }
        }

        ExcelConditionalFormattingDataBar(ExcelConditionalFormattingDataBar copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
            Uid = copy.Uid;
            LowValue = copy.LowValue;
            HighValue = copy.HighValue;
            FillColor = copy.FillColor;
            BorderColor = copy.BorderColor;
            NegativeBorderColor = copy.NegativeBorderColor;
            NegativeFillColor = copy.NegativeFillColor;
            AxisColor = copy.AxisColor;

            Border = copy.Border;
            ShowValue = copy.ShowValue;
            Gradient = copy.Gradient;
            NegativeBarBorderColorSameAsPositive = copy.NegativeBarBorderColorSameAsPositive;
            NegativeBarColorSameAsPositive = copy.NegativeBarColorSameAsPositive;
            AxisPosition = copy.AxisPosition;
        }

        internal static string NewId()
        {
            return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingDataBar(this, newWs);
        }

        /// <summary>
        /// Show value
        /// </summary>
        public bool ShowValue { get; set; } = true;

        public bool Gradient { get; set; } = true;

        public bool Border { get; set; } = false;

        public bool NegativeBarColorSameAsPositive { get; set; } = true;

        public bool NegativeBarBorderColorSameAsPositive { get; set; } = true;


        public eExcelDatabarAxisPosition AxisPosition { get; set; }

        /// <summary>
        /// Databar Low Value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue LowValue { get; internal set; }

        /// <summary>
        /// Databar High Value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue HighValue { get; internal set; }
        /// <summary>
        /// Shorthand for the Fillcolor.Color property as it is the most commonly used
        /// </summary>
        public Color Color 
        { 
            get 
            {
                if(FillColor.Color != null)
                {
                    return (Color)FillColor.Color;
                }
                else
                {
                    return Color.Empty;
                }
            } 
            set
            {
                FillColor.Color = value;
            }
        }

        public ExcelDxfColor FillColor { get; set; }
        public ExcelDxfColor BorderColor { get; set; }
        public ExcelDxfColor NegativeFillColor { get; set; }
        public ExcelDxfColor NegativeBorderColor { get; set; }
        public ExcelDxfColor AxisColor { get; set; }

        public eDatabarDirection Direction { get; set; }
    }
}
