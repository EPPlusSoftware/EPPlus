using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;
using System.Globalization;
using System;
using OfficeOpenXml.Utils.Extensions;

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

            //Excel default blue?
            Color = Color.FromArgb(int.Parse("FF638EC6", NumberStyles.HexNumber));

            var colVal = int.Parse("FFFF0000", NumberStyles.HexNumber);
            NegativeFillColor = Color.FromArgb(colVal);
            AxisColor = Color.FromArgb(colVal);
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
            Color = copy.Color;
            NegativeFillColor = copy.NegativeFillColor;
            AxisColor = copy.AxisColor;
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
        public bool ShowValue { get; set; }
        /// <summary>
        /// Databar Low Value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue LowValue { get; internal set; }

        /// <summary>
        /// Databar High Value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue HighValue { get; internal set; }
        /// <summary>
        /// The color of the databar
        /// </summary>
        public Color Color { get; set; }

        public Color FillColor { get; set; }
        public Color BorderColor { get; set; }
        public Color NegativeFillColor { get; set; }
        public Color NegativeBorderColor { get; set; }
        public Color AxisColor { get; set; }
    }
}
