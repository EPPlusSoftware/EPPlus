/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount
{
    internal class FinancialDaysEuropean_30_360 : FinancialDaysBase, IFinanicalDays
    {

        public double GetDaysBetweenDates(System.DateTime startDate, System.DateTime endDate)
        {
            var start = FinancialDayFactory.Create(startDate, DayCountBasis.Actual_Actual);
            var end = FinancialDayFactory.Create(endDate, DayCountBasis.Actual_Actual);
            return GetDaysBetweenDates(start, end, (int)DaysPerYear);
        }

        public double GetDaysBetweenDates(FinancialDay startDate, FinancialDay endDate)
        {
            if (startDate.Day == 31) startDate.Day = 30;
            if (endDate.Day == 31) endDate.Day = 30;
            return GetDaysBetweenDates(startDate, endDate, (int)DaysPerYear);
        }

        public double GetCoupdays(FinancialDay start, FinancialDay end, int frequency)
        {
            return DaysPerYear / frequency;
        }

        public double DaysPerYear { get { return 360d; } }
    }
}
