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
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingThreeColorScale : ExcelConditionalFormattingTwoColorScale,
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

        internal ExcelConditionalFormattingThreeColorScale(ExcelConditionalFormattingThreeColorScale copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
            LowValue = copy.LowValue;
            MiddleValue = copy.MiddleValue;
            HighValue = copy.HighValue;
            StopIfTrue = copy.StopIfTrue;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingThreeColorScale(this, newWs);
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
        internal override string ApplyStyleOverride(ExcelAddress address)
        {
            var range = _ws.Cells[address.Address];
            var cellValue = range.Value;
            if (cellValue.IsNumeric())
            {
                var cellValues = new List<object>();
                foreach (var cell in Address.GetAllAddresses())
                {
                    for (int i = 1; i <= cell.Rows; i++)
                    {
                        for (int j = 1; j <= cell.Columns; j++)
                        {
                            cellValues.Add(_ws.Cells[cell._fromRow + i - 1, cell._fromCol + j - 1].Value);
                        }
                    }
                }

                var values = cellValues.OrderBy(n => n);
                int index = 0;

                foreach (var value in values)
                {
                    if (value == cellValue)
                    {
                        break;
                    }
                    index++;
                }

                float percentage = (float)index / (values.Count()/2);

                var lowcol = LowValue.Color;
                var midCol = MiddleValue.Color;
                var highcol = HighValue.Color;

                var r = lowcol.R;
                var g = lowcol.G;
                var b = lowcol.B;

                var mR = midCol.R;
                var mG = midCol.G;
                var mB = midCol.B;

                var hiR = highcol.R;
                var hiG = highcol.G;
                var hiB = highcol.B;

                var originalPercent = 1.0 - percentage;
                Color newColor;

                if(originalPercent > 0)
                {
                    var absR = (int)Math.Abs(originalPercent * r - mR * percentage);
                    var absG = (int)Math.Abs(originalPercent * g - mG * percentage);
                    var absB = (int)Math.Abs(originalPercent * b - mB * percentage);
                    newColor = Color.FromArgb(1, absR, absG, absB);
                }
                else
                {
                    percentage = ((float)index - values.Count()/2) / values.Count();
                    originalPercent = 1.0 - percentage;
                    var absR = (int)Math.Abs(originalPercent * mR - hiR * percentage);
                    var absG = (int)Math.Abs(originalPercent * mG - hiG * percentage);
                    var absB = (int)Math.Abs(originalPercent * mB - hiB * percentage);
                    newColor = Color.FromArgb(1, absR, absG, absB);
                }

                return "#" + newColor.ToArgb().ToString("x8").Substring(2);
            }
            return "";
        }
    }
}
