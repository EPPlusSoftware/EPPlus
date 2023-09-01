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
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Two Colour Scale class
    /// </summary>
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
                if(double.TryParse(lowVal, out Double res))
                {
                    LowValue.Value = double.Parse(lowVal, CultureInfo.InvariantCulture);
                }
                else
                {
                    LowValue.Formula = lowVal;
                }
            }

            if (!string.IsNullOrEmpty(highVal))
            {
                if (double.TryParse(highVal, out Double res))
                {
                    HighValue.Value = double.Parse(highVal, CultureInfo.InvariantCulture);
                }
                else
                {
                    HighValue.Formula = highVal;
                }
            }
        }

        internal ExcelConditionalFormattingTwoColorScale(ExcelConditionalFormattingTwoColorScale copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
            LowValue = copy.LowValue;
            HighValue = copy.HighValue;
            StopIfTrue = copy.StopIfTrue;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingTwoColorScale(this, newWs);
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

            ReadColorAndColorSettings(xr, ref _lowValue);

            xr.Read();

            ReadColorAndColorSettings(xr, ref _highValue);

            xr.Read();
            xr.Read();
        }

        /// <summary>
        /// Internal Reading function
        /// </summary>
        /// <param name="xr"></param>
        /// <param name="colSettings"></param>
        internal void ReadColorAndColorSettings(XmlReader xr, ref ExcelConditionalFormattingColorScaleValue colSettings)
        {
            if (!string.IsNullOrEmpty(xr.GetAttribute("auto")))
            {
                colSettings.ColorSettings.Auto = xr.GetAttribute("auto") == "1" ? true : false;
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("theme")))
            {
                colSettings.ColorSettings.Theme = (eThemeSchemeColor)int.Parse(xr.GetAttribute("theme"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("indexed")))
            {
                colSettings.ColorSettings.Index = int.Parse(xr.GetAttribute("indexed"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("rgb")))
            {
                colSettings.ColorSettings.Color = (ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb")));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("tint")))
            {
                colSettings.ColorSettings.Tint = double.Parse(xr.GetAttribute("tint"));
            }
        }

        internal ExcelConditionalFormattingColorScaleValue _lowValue;
        internal ExcelConditionalFormattingColorScaleValue _highValue;


        /// <summary>
        /// Low Value for Two Color Scale Object Value
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue LowValue
        {
            get { return _lowValue; }
            set { _lowValue = value; }
        }

        /// <summary>
        /// High Value for Two Color Scale Object Value
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue HighValue
        {
            get { return _highValue; }
            set { _highValue = value; }
        }
    }
}
