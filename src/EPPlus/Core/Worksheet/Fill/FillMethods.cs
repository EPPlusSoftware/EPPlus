/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    08/11/2021         EPPlus Software AB       EPPlus 5.8
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet.Fill
{
    internal class FillMethods
    {
        internal static void FillNumber(ExcelWorksheet worksheet, int fromRow, int toRow, int fromCol, int toCol, FillNumberParams options)
        {
            object startValue;
            GetStartCell(options, fromRow, toRow, fromCol, toCol, out int startRow, out int startCol);
            if (options.StartValue.HasValue)
            {
                startValue = options.StartValue;
                worksheet.SetValue(startRow, startCol, startValue);
            }
            else
            {
                startValue = worksheet.GetValue(startRow, startCol);
            }
            var value = ConvertUtil.GetValueDouble(startValue, true, true);

            SkipFirstCell(ref fromRow, ref fromCol, ref toRow, ref toCol, options);

            int r = startRow, c = startCol;
            while (GetNextCell(options, fromRow, toRow, fromCol, toCol, ref r, ref c))
            {
                FillCellNumber(worksheet, options, ref value, r, c);
            }
        }

        private static void FillCellNumber(ExcelWorksheet worksheet, FillNumberParams options, ref double value, int r, int c)
        {
            if (double.IsNaN(value))
            {
                worksheet.SetValue(r, c, null);
            }
            else
            {
                if (options.CalculationMethod == eCalculationMethod.Add)
                {
                    value += options.StepValue;
                }
                else
                {
                    value *= options.StepValue;
                }
                if (options.EndValue.HasValue && options.EndValue.Value < value)
                {
                    worksheet.SetValue(r, c, null);
                }
                else
                {
                    worksheet.SetValue(r, c, value);
                }
            }
        }

        internal static void FillDateTime(ExcelWorksheet worksheet, int fromRow, int toRow, int fromCol, int toCol, FillDateParams options)
        {
            object startValue;
            GetStartCell(options, fromRow, toRow, fromCol, toCol, out int startRow, out int startCol);
            if (options.StartValue.HasValue)
            {
                startValue = options.StartValue;
                worksheet.SetValue(startRow, startCol, startValue);
            }
            else
            {
                startValue = worksheet.GetValue(startRow, startCol);
            }
            //if (options.Direction == eFillDirection.Column) fromRow++; else fromCol++;
            SkipFirstCell(ref fromRow, ref fromCol, ref toRow, ref toCol, options);

            var value = ConvertUtil.GetValueDate(startValue);
            var isLastDayOfMonth = value.HasValue && value.Value.Month != value.Value.AddDays(1).Month;

            int r = startRow, c = startCol;
            while (GetNextCell(options, fromRow, toRow, fromCol, toCol, ref r, ref c))
            {
                FillCellDate(worksheet, options, ref value, isLastDayOfMonth, c, r);
            }
        }

        private static void FillCellDate(ExcelWorksheet worksheet, FillDateParams options, ref DateTime? value, bool isLastDayOfMonth, int c, int r)
        {
            if (value.HasValue)
            {
                switch (options.DateUnit)
                {
                    case eDateTimeUnit.Year:
                        value = value.Value.AddYears(options.StepValue);
                        break;
                    case eDateTimeUnit.Month:
                        if (isLastDayOfMonth)
                        {
                            value = value.Value.AddMonths(options.StepValue + 1);
                            value = value.Value.AddDays(-value.Value.Day);
                        }
                        else
                        {
                            value = value.Value.AddMonths(options.StepValue);
                        }
                        break;
                    case eDateTimeUnit.Week:
                        value = value.Value.AddDays(options.StepValue * 7);
                        break;
                    case eDateTimeUnit.Day:
                        value = value.Value.AddDays(options.StepValue);
                        break;
                    case eDateTimeUnit.Hour:
                        value = value.Value.AddHours(options.StepValue);
                        break;
                    case eDateTimeUnit.Minute:
                        value = value.Value.AddMinutes(options.StepValue);
                        break;
                    case eDateTimeUnit.Second:
                        value = value.Value.AddSeconds(options.StepValue);
                        break;
                    case eDateTimeUnit.Ticks:
                        value = value.Value.AddTicks(options.StepValue);
                        break;
                }
                DateTime d;
                if (options.WeekdaysOnly)
                {
                    d = GetWeekday(value.Value, options.HolidayCalendar);
                }
                else
                {
                    d = value.Value;
                }

                if (options.EndValue == null || value <= options.EndValue)
                {
                    worksheet.SetValue(r, c, d);
                }
                else
                {
                    worksheet.SetValue(r, c, null);
                }
            }
            else
            {
                worksheet.SetValue(r, c, null);
            }
        }

        internal static void FillList<T>(ExcelWorksheet worksheet, int fromRow, int toRow, int fromCol, int toCol,IEnumerable<T> enumList, FillListParams options)
        {
            var list = enumList.ToList();

            if (list.Count==0)
            {
                worksheet.Cells[fromRow, fromCol, toRow, toCol].Clear();
                return;
            }

            if (options.StartIndex<0 || options.StartIndex>=list.Count)
            {
                throw new InvalidOperationException("StartIndex must be within the list");
            }

            GetStartCell(options, fromRow, toRow, fromCol, toCol, out int startRow, out int startCol);
            var ix = options.StartIndex;
            worksheet.SetValue(startRow, startCol, list[ix++]);
            SkipFirstCell(ref fromRow, ref fromCol, ref toRow, ref toCol, options);

            int r = startRow, c = startCol;
            while (GetNextCell(options, fromRow, toRow, fromCol, toCol, ref r, ref c))
            {
                if (ix == list.Count) ix = 0;
                    worksheet.SetValue(r, c, list[ix++]);
            }
        }
        private static bool GetNextCell(FillParams options, int fromRow, int toRow, int fromCol, int toCol, ref int r, ref int c)
        {
            switch (options.StartPosition)
            {
                case eFillStartPosition.TopLeft:
                    c++;
                    if (c > toCol)
                    {
                        r++;
                        if (r > toRow)
                        {
                            return false;
                        }
                        c = fromCol;
                    }
                    break;
                default:
                    c--;
                    if (c < fromCol)
                    {
                        r--;
                        if (r < fromRow)
                        {
                            return false;
                        }
                        c = toCol;
                    }
                    break;
            }
            return true;
        }

        private static void GetStartCell(FillParams options, int fromRow, int toRow, int fromCol, int toCol, out int startRow, out int startCol)
        {
            switch (options.StartPosition)
            {
                case eFillStartPosition.TopLeft:
                    startRow = fromRow;
                    startCol = fromCol;
                    break;
                default:
                    startRow = toRow;
                    startCol = toCol;
                    break;
            }
        }
        private static void SkipFirstCell(ref int fromRow, ref int fromCol, ref int toRow, ref int toCol, FillParams options)
        {
            if (options.StartPosition == eFillStartPosition.TopLeft)
            {
                if (options.Direction == eFillDirection.Column)
                    fromRow++;
                else
                    fromCol++;
            }
            else
            {
                if (options.Direction == eFillDirection.Column)
                    toRow--;
                else
                    toCol--;
            }

        }

        static DateTime GetWeekday(DateTime value, HashSet<DateTime> holyDays)
        {
            while (value.DayOfWeek == DayOfWeek.Saturday || value.DayOfWeek == DayOfWeek.Sunday || holyDays.Contains(value))
            {
                value = value.AddDays(-1);
            }
            return value;
        }
    }    
}
