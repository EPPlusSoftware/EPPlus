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
using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

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
