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
    internal class FinancialDays_Actual_360 : FinancialDaysBase, IFinanicalDays
    {

        public double GetDaysBetweenDates(DateTime startDate, DateTime endDate)
        {
            var start = FinancialDayFactory.Create(startDate, DayCountBasis.Actual_360);
            var end = FinancialDayFactory.Create(endDate, DayCountBasis.Actual_360);
            return GetDaysBetweenDates(start, end, 360);
        }

        public double GetDaysBetweenDates(FinancialDay startDate, FinancialDay endDate)
        {
            return GetDaysBetweenDates(startDate, endDate, (int)DaysPerYear);
        }

        protected override double GetDaysBetweenDates(FinancialDay start, FinancialDay end, int basis)
        {
            return end.ToDateTime().Subtract(start.ToDateTime()).TotalDays;
        }

        public double GetCoupdays(FinancialDay start, FinancialDay end, int frequency)
        {
            return DaysPerYear / frequency;
        }

        public double DaysPerYear { get { return 360d; } }
    }
}
