using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.ConditionalFormatting
{

    public class ExcelConditionalFormattingTwoColorScale : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTwoColorScale
    {
        internal ExcelConditionalFormattingTwoColorScale(
        ExcelAddress address,
        int priority,
        ExcelWorksheet ws) 
        : base(
            eExcelConditionalFormattingRuleType.TwoColorScale, 
            address, 
            priority, 
            ws)
        {
            LowValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Min,
                ExcelConditionalFormattingConstants.Colors.CfvoLowValue, 
                priority);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Max,
                ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
                priority);
        }

        internal ExcelConditionalFormattingTwoColorScale(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet,
            bool stopIfTrue,
            eExcelConditionalFormattingValueObjectType? low, 
            eExcelConditionalFormattingValueObjectType? high,
            string lowVal,
            string highVal,
            XmlReader xr) : base (
                eExcelConditionalFormattingRuleType.TwoColorScale, 
                address, 
                priority, 
                worksheet)
        {
            StopIfTrue = stopIfTrue;
            SetValues(low, high, lowVal, highVal);
            ReadColors(xr);
        }

        internal void SetValues(
            eExcelConditionalFormattingValueObjectType? low, 
            eExcelConditionalFormattingValueObjectType? high,
            string lowVal = "",
            string highVal = "",
            string middleVal = "",
            eExcelConditionalFormattingValueObjectType? middle = null)
        {
            LowValue = new ExcelConditionalFormattingColorScaleValue(
            low,
            ExcelConditionalFormattingConstants.Colors.CfvoLowValue,
            Priority);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
            high,
            ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
            Priority);

            if (!string.IsNullOrEmpty(lowVal))
            {
                if (low == eExcelConditionalFormattingValueObjectType.Formula)
                {
                    LowValue.Formula = lowVal;
                }
                else
                {
                    LowValue.Value = double.Parse(lowVal, CultureInfo.InvariantCulture);
                }
            }

            if (!string.IsNullOrEmpty(highVal)) 
            {
                if (high == eExcelConditionalFormattingValueObjectType.Formula)
                {
                    HighValue.Formula = highVal;
                }
                else
                {
                    HighValue.Value = double.Parse(highVal, CultureInfo.InvariantCulture);
                }
            }
        }

        internal ExcelConditionalFormattingTwoColorScale(ExcelConditionalFormattingTwoColorScale copy) : base(copy)
        {
            LowValue = copy.LowValue;
            HighValue = copy.HighValue;
            StopIfTrue = copy.StopIfTrue;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingTwoColorScale(this);
        }

        internal override bool IsExtLst
        {
            get
            {
                if (ExcelAddressBase.RefersToOtherWorksheet(LowValue.Formula, _ws.Name) || 
                    ExcelAddressBase.RefersToOtherWorksheet(HighValue.Formula, _ws.Name))
                {
                    return true;
                }

                return base.IsExtLst;
            } 

        }

        internal virtual void ReadColors(XmlReader xr)
        {
            Type = eExcelConditionalFormattingRuleType.TwoColorScale;
            LowValue.Color = (ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb")));

            xr.Read();

            HighValue.Color = (ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb")));

            xr.Read();
            xr.Read();
        }

        /// <summary>
        /// Low Value for Two Color Scale Object Value
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue LowValue
        {
            get;
            set;
        }

        /// <summary>
        /// High Value for Two Color Scale Object Value
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue HighValue
        {
            get;
            set;
        }
    }
}
