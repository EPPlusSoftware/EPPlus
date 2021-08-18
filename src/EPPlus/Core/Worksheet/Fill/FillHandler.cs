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
    internal class FillHandler
    {
        internal static void FillNumbers(ExcelWorksheet worksheet, int fromRow, int toRow, int fromCol, int toCol, FillNumberParams options)
        {
            object startValue;
            if (options.StartValue.HasValue)
            {
                startValue = options.StartValue;
                worksheet.SetValue(fromRow, fromCol, startValue);
            }
            else
            {
                startValue = worksheet.GetValue(fromRow, fromCol);
            }
            if (options.Direction == eFillDirection.Column) fromRow++; else fromCol++;
            var value = ConvertUtil.GetValueDouble(startValue, true, true);
            for (int c = fromCol; c <= toCol; c++)
            {
                for (int r = fromRow; r <= toRow; r++)
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
            }
        }
        internal static void FillDates(ExcelWorksheet worksheet, int fromRow, int toRow, int fromCol, int toCol, FillDateParams options)
        {
            object startValue;
            if (options.StartValue.HasValue)
            {
                startValue = options.StartValue;
                worksheet.SetValue(fromRow, fromCol, startValue);
            }
            else
            {
                startValue = worksheet.GetValue(fromRow, fromCol);
            }
            if (options.Direction == eFillDirection.Column) fromRow++; else fromCol++;
            var value = ConvertUtil.GetValueDate(startValue);
            var isLastDayOfMonth = value.HasValue && value.Value.Month != value.Value.AddDays(1).Month;
            for (int c = fromCol; c <= toCol; c++)
            {
                for (int r = fromRow; r <= toRow; r++)
                {
                    if (value.HasValue)
                    {
                        switch (options.DateUnit)
                        {
                            case eDateUnit.Year:
                                value = value.Value.AddYears(options.StepValue);
                                break;
                            case eDateUnit.Month:
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
                            case eDateUnit.Week:
                                value = value.Value.AddDays(options.StepValue * 7);
                                break;
                            case eDateUnit.Day:
                                value = value.Value.AddDays(options.StepValue);
                                break;
                            case eDateUnit.Hour:
                                value = value.Value.AddHours(options.StepValue);
                                break;
                            case eDateUnit.Minute:
                                value = value.Value.AddMinutes(options.StepValue);
                                break;
                            case eDateUnit.Second:
                                value = value.Value.AddSeconds(options.StepValue);
                                break;
                            case eDateUnit.Ticks:
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

                        if (options.EndValue==null || value <= options.EndValue)
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

            var ix = options.StartIndex;
            worksheet.SetValue(fromRow, fromCol, list[ix++]);
            if (options.Direction == eFillDirection.Column) fromRow++; else fromCol++;

            for (int c = fromCol; c <= toCol; c++)
            {
                for (int r = fromRow; r <= toRow; r++)
                {
                    if (ix == list.Count) ix = 0;
                    worksheet.SetValue(r, c, list[ix++]);
                }
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
