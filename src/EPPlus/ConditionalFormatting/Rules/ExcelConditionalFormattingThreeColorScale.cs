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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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
                double midPoint = 0;
                double average = 0;
                int count = 0;
                foreach (var cell in Address.GetAllAddresses())
                {
                    for (int i = 1; i <= cell.Rows; i++)
                    {
                        for (int j = 1; j <= cell.Columns; j++)
                        {
                            cellValues.Add(_ws.Cells[cell._fromRow + i - 1, cell._fromCol + j - 1].Value);
                            average += Convert.ToDouble(_ws.Cells[cell._fromRow + i - 1, cell._fromCol + j - 1].Value);
                            count++;
                        }
                    }
                }

                average = average / count;

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

                var highest = Convert.ToDouble(values.Last());
                var lowest = Convert.ToDouble(values.First());
                //midPoint = (highest + lowest) * 0.5;
                var realValue = Convert.ToDouble(cellValue);

                Color newColor;

                if (realValue < average)
                {
                    newColor = CalculateNumberedGradient(realValue - lowest, average - lowest, LowValue.Color, MiddleValue.Color);
                }
                else
                {
                    newColor = CalculateNumberedGradient(realValue - average, highest - average, MiddleValue.Color, HighValue.Color);
                    //newColor = CalculateNumberedGradient(realValue, highest, MiddleValue.Color, HighValue.Color);
                }



                //205,92,92) 
                //(208, 102, 100);
                //(211, 113, 108)
                //(214, 124, 116)
                //(217, 135, 124)
                //(220, 146, 132)
                //(223, 157, 140)
                //(226, 168, 148)
                //(229, 179, 157)
                //(232, 190, 165)
                //(236, 200, 173)
                //(239, 211, 181)
                //(242, 222, 189)
                //(245, 233, 197)
                //(248, 244, 205)

                //int halfTotal = (int)Math.Round(((double)values.Count() * 0.5d));

                ////Note: index is/must be 0 based so first cell == LowValue.Color
                ////Therefore the midpoint cell is halftotal-1 as at e.g. row 15 out of 30 index = 14
                //int midPoint = halfTotal - 1;

                //int midPoint = halfTotal - 1;
                //Color newColor;
                //if (Convert.ToDouble(cellValue) < average)
                //{
                //    newColor = CalculateNumberedGradient(Convert.ToDouble(cellValue), average, LowValue.Color, MiddleValue.Color);
                //}
                //else
                //{
                //    newColor = CalculateNumberedGradient(Convert.ToDouble(cellValue) -average, average, MiddleValue.Color, HighValue.Color);
                //}



                //if (index < (average-1))
                //{
                //    newColor = CalculateNumberedGradient(index, average, LowValue.Color, MiddleValue.Color);
                //}
                //else
                //{
                //    if(index != count-1)
                //    {
                //        newColor = CalculateNumberedGradient(index - average, average, MiddleValue.Color, HighValue.Color);
                //    }
                //    else
                //    {
                //        newColor = HighValue.Color;
                //    }
                //}

                //if (index < average)
                //{
                //    newColor = CalculateNumberedGradient(index, average, LowValue.Color, MiddleValue.Color);
                //}
                //else
                //{
                //    newColor = CalculateNumberedGradient((double)index - average, average, MiddleValue.Color, HighValue.Color);
                //}

                return "#" + newColor.ToArgb().ToString("x8").Substring(2);
            }
            return "";
        }
    }
}
