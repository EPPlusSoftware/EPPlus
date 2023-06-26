using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingThreeColorScale : ExcelConditionalFormattingTwoColorScale,
    IExcelConditionalFormattingThreeColorScale
    {

        private Color tempColor;

        internal ExcelConditionalFormattingThreeColorScale(ExcelAddress address, int priority, ExcelWorksheet ws)
            : base(address, priority, ws)
        {
            MiddleValue = new ExcelConditionalFormattingColorScaleValue(
            eExcelConditionalFormattingValueObjectType.Percentile,
            ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue,
            priority);

            Type = eExcelConditionalFormattingRuleType.ThreeColorScale;

            MiddleValue.Value = 50;
        }

        internal ExcelConditionalFormattingThreeColorScale(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet,
        bool stopIfTrue,
        eExcelConditionalFormattingValueObjectType? low,
        eExcelConditionalFormattingValueObjectType? middle,
        eExcelConditionalFormattingValueObjectType? high,
        string lowVal,
        string middleVal,
        string highVal,
        XmlReader xr) : base(
            address,
            priority,
            worksheet,
            stopIfTrue, 
            low, 
            high, 
            lowVal, 
            highVal, 
            xr)
        {
            MiddleValue = new ExcelConditionalFormattingColorScaleValue(
            middle,
            ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue,
            Priority);

            MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;

            if (!string.IsNullOrEmpty(middleVal))
            {
                if (middle == eExcelConditionalFormattingValueObjectType.Formula)
                {
                    MiddleValue.Formula = middleVal;
                }
                else
                {
                    MiddleValue.Value = double.Parse(middleVal, CultureInfo.InvariantCulture);
                }
            }

            MiddleValue.Color = tempColor;
        }

        internal ExcelConditionalFormattingThreeColorScale(ExcelConditionalFormattingThreeColorScale copy) : base(copy)
        {
            LowValue = copy.LowValue;
            MiddleValue = copy.MiddleValue;
            HighValue = copy.HighValue;
            StopIfTrue = copy.StopIfTrue;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingThreeColorScale(this);
        }

        internal override bool IsExtLst
        {
            get
            {
                if (ExcelAddressBase.RefersToOtherWorksheet(MiddleValue.Formula, _ws.Name))
                {
                    return true;
                }

                return base.IsExtLst;
            }

        }

        internal override void ReadColors(XmlReader xr)
        {
            //we don't call base as the order of nodes are different. Second node is middle.

            Type = eExcelConditionalFormattingRuleType.ThreeColorScale;
            string test = xr.GetAttribute("rgb");

            LowValue.Color = ExcelConditionalFormattingHelper.ConvertFromColorCode(test);

            xr.Read();

            tempColor = ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb"));

            xr.Read();

            HighValue.Color = ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb"));

            xr.Read();
            xr.Read();
        }

        public ExcelConditionalFormattingColorScaleValue MiddleValue
        {
            get;
            set;
        }
    }
}
