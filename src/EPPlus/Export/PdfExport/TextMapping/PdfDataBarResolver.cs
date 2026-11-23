/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Drawing;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfDataBarResolver
    {
        internal static ExcelConditionalFormattingDataBar GetDataBarRule(ExcelRangeBase cell)
        {
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf == null || cf.Count == 0)
            {
                return null;
            }
            foreach (var rule in cf.OrderBy(r => r.Priority))
            {
                if (!rule.ShouldApplyToCell(cell))
                {
                    if (rule.StopIfTrue)
                    {
                        break;
                    }
                    continue;
                }
                if (rule is ExcelConditionalFormattingDataBar db)
                {
                    return db;
                }
                if (rule.StopIfTrue)
                {
                    break;
                }
            }
            return null;
        }

        internal static void GetDataBarStyle(ExcelRange cell, PdfCellStyle cellStyle, ExcelConditionalFormattingDataBar dataBarRule)
        {
            if (dataBarRule != null)
            {
                double low = dataBarRule.lowest, high = dataBarRule.highest;
                double axisFraction;
                if (dataBarRule.AxisPosition == eExcelDatabarAxisPosition.None) axisFraction = 0d;
                else if (dataBarRule.AxisPosition == eExcelDatabarAxisPosition.Middle) axisFraction = 0.5d;
                else if (low >= 0d) axisFraction = 0d;                 // all positive: axis at left
                else if (high <= 0d) axisFraction = 1d;                 // all negative: axis at right
                else axisFraction = Math.Abs(low) / (high - low);   // mixed

                // None shows negatives in the positive direction, so only flag negatives when there's an axis.
                bool isNeg = dataBarRule.AxisPosition != eExcelDatabarAxisPosition.None
                             && ConvertUtil.GetValueDouble(cell.Value) < 0d;

                var fillCol = (isNeg ? dataBarRule.NegativeFillColor : dataBarRule.FillColor).GetColorAsColor();
                Color? border = dataBarRule.Border
                    ? (isNeg ? dataBarRule.NegativeBorderColor : dataBarRule.BorderColor).GetColorAsColor()
                    : (Color?)null;
                Color? axis = (axisFraction > 0d && axisFraction < 1d)
                    ? (dataBarRule.AxisColor != null && dataBarRule.AxisColor.HasValue
                          ? dataBarRule.AxisColor.GetColorAsColor()
                          : Color.FromArgb(0x14, 0x09, 0x04))          // Excel's default dark axis
                    : (Color?)null;

                cellStyle.DataBar = new PdfDataBar
                {
                    Percentage = dataBarRule.GetPercentageAtCell(cell),
                    IsNegative = isNeg,
                    AxisFraction = axisFraction,
                    FillColor = fillCol,
                    BorderColor = border,
                    AxisColor = axis,
                    Gradient = dataBarRule.Gradient,
                    IsRightToLeft = dataBarRule.Direction == eDatabarDirection.RightToLeft
                };
            }
        }

        internal static IExcelConditionalFormattingIconSet GetIconSetRule(ExcelRangeBase cell)
        {
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf == null || cf.Count == 0) return null;
            foreach (var rule in cf.OrderBy(r => r.Priority))
            {
                if (!rule.ShouldApplyToCell(cell)) { if (rule.StopIfTrue) break; continue; }
                if (rule is IExcelConditionalFormattingIconSet iconSet) return iconSet;
                if (rule.StopIfTrue) break;
            }
            return null;
        }
    }
}
