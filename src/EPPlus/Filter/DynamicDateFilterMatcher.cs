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
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Filter
{
    internal static class DynamicDateFilterMatcher
    {
        internal static bool Match(ExcelDynamicFilterColumn column, DateTime? value)
        {
            switch(column.Type)
            {
                case eDynamicFilterType.Yesterday:
                    return value.Value.Date == DateTime.Today.AddDays(-1);
                case eDynamicFilterType.Today:
                    return value.Value.Date == DateTime.Today;
                case eDynamicFilterType.Tomorrow:
                    return value.Value.Date == DateTime.Today.AddDays(1);
                case eDynamicFilterType.M1:
                    return value.Value.Month == 1;
                case eDynamicFilterType.M2:
                    return value.Value.Month == 2;
                case eDynamicFilterType.M3:
                    return value.Value.Month == 3;
                case eDynamicFilterType.M4:
                    return value.Value.Month == 4;
                case eDynamicFilterType.M5:
                    return value.Value.Month == 5;
                case eDynamicFilterType.M6:
                    return value.Value.Month == 6;
                case eDynamicFilterType.M7:
                    return value.Value.Month == 7;
                case eDynamicFilterType.M8:
                    return value.Value.Month == 8;
                case eDynamicFilterType.M9:
                    return value.Value.Month == 9;
                case eDynamicFilterType.M10:
                    return value.Value.Month == 10;
                case eDynamicFilterType.M11:
                    return value.Value.Month == 11;
                case eDynamicFilterType.M12:
                    return value.Value.Month == 12;
                case eDynamicFilterType.Q1:
                    return value.Value.Month >= 1 && value.Value.Month <= 3;
                case eDynamicFilterType.Q2:
                    return value.Value.Month >= 4 && value.Value.Month <= 6;
                case eDynamicFilterType.Q3:
                    return value.Value.Month >= 7 && value.Value.Month <= 9;
                case eDynamicFilterType.Q4:
                    return value.Value.Month >= 10 && value.Value.Month <= 12;                
                default:
                    var v = value.Value.ToOADate();
                    return v >= column.Value && v <= column.MaxValue;
            }
        }
        internal static void SetMatchDates(ExcelDynamicFilterColumn column)
        {
            switch (column.Type)
            {
                case eDynamicFilterType.Yesterday:
                    SetDay(column, DateTime.Today.AddDays(-1));
                    break;
                case eDynamicFilterType.Today:
                    SetDay(column, DateTime.Today);
                    break;
                case eDynamicFilterType.Tomorrow:
                    SetDay(column, DateTime.Today.AddDays(1));
                    break;
                case eDynamicFilterType.LastWeek:
                    SetWeek(column, DateTime.Today.AddDays(-7));
                    break;
                case eDynamicFilterType.ThisWeek:
                    SetWeek(column, DateTime.Today);
                    break;
                case eDynamicFilterType.NextWeek:
                    SetWeek(column, DateTime.Today.AddDays(7));
                    break;
                case eDynamicFilterType.LastMonth:
                    SetFullMonth(column, DateTime.Today.AddMonths(-1));
                    break;
                case eDynamicFilterType.ThisMonth:
                    SetFullMonth(column, DateTime.Today);
                    break;
                case eDynamicFilterType.NextMonth:
                    SetFullMonth(column, DateTime.Today.AddMonths(1));
                    break;
                case eDynamicFilterType.LastQuarter:
                    SetFullQuarter(column, DateTime.Today.AddMonths(-3));
                    break;
                case eDynamicFilterType.ThisQuarter:
                    SetFullQuarter(column, DateTime.Today);
                    break;
                case eDynamicFilterType.NextQuarter:
                    SetFullQuarter(column, DateTime.Today.AddMonths(3));
                    break;
                case eDynamicFilterType.LastYear:
                    SetFullYear(column, DateTime.Today.Year-1);
                    break;
                case eDynamicFilterType.ThisYear:
                    SetFullYear(column, DateTime.Today.Year);
                    break;
                case eDynamicFilterType.NextYear:
                    SetFullYear(column, DateTime.Today.Year + 1);
                    break;
                case eDynamicFilterType.YearToDate:
                    SetYearToDate(column);
                    break;
                default:
                    SetFixed(column);
                    break;
            }
        }

        private static void SetFixed(ExcelDynamicFilterColumn column)
        {
                column.Value = null;
                column.MaxValue = null;
        }

        private static void SetDay(ExcelDynamicFilterColumn column, DateTime dt)
        {
            column.Value = dt.ToOADate();
            column.MaxValue = dt.AddDays(1).ToOADate();
        }

        private static void SetYearToDate(ExcelDynamicFilterColumn column)
        {
            column.Value = new DateTime(DateTime.Today.Year, 1, 1).ToOADate();
            column.MaxValue = DateTime.Today.ToOADate();
        }

        private static void SetFullQuarter(ExcelDynamicFilterColumn column, DateTime dt)
        {
            var quarter = ((dt.Month - (dt.Month - 1) % 3) + 1) / 3;
            var minDate = new DateTime(dt.Year, (quarter * 3) + 1, 1);
            column.Value = minDate.ToOADate();
            column.MaxValue = minDate.AddMonths(3).AddDays(-1).ToOADate();
        }

        private static void SetFullYear(ExcelDynamicFilterColumn column, int year)
        {
            column.Value = new DateTime(year, 1, 1).ToOADate();
            column.MaxValue = new DateTime(year, 12, 31).ToOADate();
        }

        private static void SetFullMonth(ExcelDynamicFilterColumn column, DateTime dt)
        {
            var minValue = new DateTime(dt.Year, dt.Month, 1);
            column.Value = minValue.ToOADate();
            column.MaxValue = minValue.AddMonths(1).AddDays(-1).ToOADate();
        }
        private static void SetWeek(ExcelDynamicFilterColumn column, DateTime dt)
        {
            while(dt.DayOfWeek!=DayOfWeek.Sunday)
            {
                dt = dt.AddDays(-1);
            }
            column.Value = dt.ToOADate();
            column.MaxValue = dt.AddDays(6).ToOADate();
        }
    }
}
