using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.RichData.Types;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingThreeColorScale : ExcelConditionalFormattingTwoColorScale,
    IExcelConditionalFormattingThreeColorScale
    {
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
            if(MiddleValue == null)
            {
                MiddleValue = new ExcelConditionalFormattingColorScaleValue(
                middle,
                ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue,
                Priority);
            }

            if (!string.IsNullOrEmpty(middleVal))
            {
                MiddleValue.Type = (eExcelConditionalFormattingValueObjectType)middle;

                if (double.TryParse(middleVal, out Double res))
                {
                    MiddleValue.Value = double.Parse(middleVal, CultureInfo.InvariantCulture);
                }
                else
                {
                    MiddleValue.Formula = middleVal;
                }
            }
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

            ReadColorAndColorSettings(xr, ref _lowValue);

            xr.Read();

            MiddleValue = new ExcelConditionalFormattingColorScaleValue(
               eExcelConditionalFormattingValueObjectType.Percentile,
               ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue,
               Priority);

            ReadColorAndColorSettings(xr, ref _middleValue);

            xr.Read();

            ReadColorAndColorSettings(xr, ref _highValue);

            xr.Read();
            xr.Read();
        }

        ExcelConditionalFormattingColorScaleValue _middleValue;

        /// <summary>
        /// The middle value.
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue MiddleValue
        {
            get { return _middleValue; }
            set { _middleValue = value; }
        }
    }
}
