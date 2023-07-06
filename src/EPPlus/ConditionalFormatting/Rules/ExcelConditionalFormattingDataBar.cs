using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;
using System.Globalization;
using System;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style;
using static OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingConstants;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingDataBar : ExcelConditionalFormattingRule,
        IExcelConditionalFormattingDataBarGroup
    {
        internal string Uid { get; set; }

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

            //var colVal = int.Parse("FFFF0000", NumberStyles.HexNumber);
            //NegativeFillColor.Color = Color.FromArgb(colVal);
            //NegativeBorderColor.Color = Color.FromArgb(colVal);

            //AxisColor.Color = Color.FromArgb(colVal);

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
                HighValue.Value = Double.Parse(xr.GetAttribute("val"));
            }

            xr.Read();
            var lowType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>().Value;
            LowValue = new ExcelConditionalFormattingIconDataBarValue(lowType, eExcelConditionalFormattingRuleType.DataBar);

            if (!string.IsNullOrEmpty(xr.GetAttribute("val")))
            {
                LowValue.Value = Double.Parse(xr.GetAttribute("val"));
            }

            xr.Read();

            InitalizeDxfColours();

            var colVal = int.Parse(xr.GetAttribute("rgb"),NumberStyles.HexNumber);
            Color = Color.FromArgb(colVal);
            //Correct the alpha
            Color = Color.FromArgb(255, Color);

            //enter databar exit node ->(local) extLst -> ext -> id
            xr.Read();
            xr.Read();
            xr.Read();
            xr.Read();

            Uid = xr.ReadString();

            // /ext -> /extLst
            xr.Read();
            xr.Read();
            xr.Read();
        }

        ExcelConditionalFormattingDataBar(ExcelConditionalFormattingDataBar copy) : base(copy)
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

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingDataBar(this);
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
    }
}
