using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;
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

            ReadColorAndColorSettings(xr, ref _lowValue);

            xr.Read();

            ReadColorAndColorSettings(xr, ref _highValue);

            xr.Read();
            xr.Read();
        }

        protected void ReadColorAndColorSettings(XmlReader xr, ref ExcelConditionalFormattingColorScaleValue colSettings)
        {
            if (!string.IsNullOrEmpty(xr.GetAttribute("auto")))
            {
                colSettings.ColorSettings.Auto = xr.GetAttribute("auto") == "1" ? true : false;
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("theme")))
            {
                colSettings.ColorSettings.Theme = (eThemeSchemeColor)int.Parse(xr.GetAttribute("theme"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("index")))
            {
                colSettings.ColorSettings.Index = int.Parse(xr.GetAttribute("index"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("rgb")))
            {
                colSettings.ColorSettings.Color = (ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb")));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("tint")))
            {
                colSettings.ColorSettings.Tint = double.Parse(xr.GetAttribute("tint"));
            }

            //return value;
        }

        protected ExcelConditionalFormattingColorScaleValue _lowValue;
        protected ExcelConditionalFormattingColorScaleValue _highValue;


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
