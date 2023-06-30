﻿/*************************************************************************************************
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
    internal interface IFinanicalDays
    {
        double GetDaysBetweenDates(DateTime startDate, DateTime endDate);

        double GetDaysBetweenDates(FinancialDay startDate, FinancialDay endDate);

        double GetDaysBetweenDates(FinancialDay startDate, FinancialDay endDate, bool returnZeroIfNegative);

        IEnumerable<FinancialPeriod> GetCouponPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency);

        IEnumerable<FinancialPeriod> GetCalendarYearPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency);
        IEnumerable<FinancialPeriod> GetCalendarYearPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency, int additionalPeriods);

        FinancialPeriod GetCouponPeriod(FinancialDay settlementDate, FinancialDay maturityDate, int frequency);

        int GetNumberOfCouponPeriods(FinancialDay settlementDate, FinancialDay maturityDate, int frequency);

        double GetCoupdays(FinancialDay startDate, FinancialDay endDate, int frequency);

        double DaysPerYear { get; }
    }
}
