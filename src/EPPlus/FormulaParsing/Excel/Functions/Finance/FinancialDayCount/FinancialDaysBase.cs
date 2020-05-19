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
    internal abstract class FinancialDaysBase
    {
        public FinancialPeriod GetCouponPeriod(FinancialDay settlementDay, FinancialDay maturityDay, int frequency)
        {
            var periods = new List<FinancialPeriod>();
            var settlementDate = settlementDay.ToDateTime();
            var tmpDay = maturityDay;
            var lastDay = tmpDay;

            while (tmpDay.ToDateTime() > settlementDate)
            {
                switch (frequency)
                {
                    case 1:
                        tmpDay = tmpDay.SubtractYears(1);
                        break;
                    case 2:
                        tmpDay = tmpDay.SubtractMonths(6, maturityDay.Day);
                        break;
                    case 4:
                        tmpDay = tmpDay.SubtractMonths(3, maturityDay.Day);
                        break;
                    default:
                        throw new ArgumentException("frequency");
                }
                if(tmpDay > settlementDay) lastDay = tmpDay;
            }
            return new FinancialPeriod(tmpDay, lastDay);
        }

        public int GetNumberOfCouponPeriods(FinancialDay settlementDay, FinancialDay maturityDay, int frequency)
        {
            var settlementDate = settlementDay.ToDateTime();
            var tmpDay = maturityDay;
            var lastDay = tmpDay;
            var nPeriods = 0;
            while (tmpDay.ToDateTime() > settlementDate)
            {
                switch (frequency)
                {
                    case 1:
                        tmpDay = tmpDay.SubtractYears(1);
                        break;
                    case 2:
                        tmpDay = tmpDay.SubtractMonths(6, maturityDay.Day);
                        break;
                    case 4:
                        tmpDay = tmpDay.SubtractMonths(3, maturityDay.Day);
                        break;
                    default:
                        throw new ArgumentException("frequency");
                }
                nPeriods++;
            }
            return nPeriods;
        }

        protected virtual double GetDaysBetweenDates(FinancialDay start, FinancialDay end, int basis)
        {
            return (basis * (end.Year - start.Year) + 30 * (end.Month - start.Month) + (end.Day - start.Day));
        }

        protected double ActualDaysInLeapYear(FinancialDay start, FinancialDay end)
        {
            var daysInLeapYear = 0d;
            for(var year = start.Year; year <= end.Year; year++)
            {
                if (System.DateTime.IsLeapYear(year))
                {
                    if(year == start.Year)
                    {
                        daysInLeapYear += new System.DateTime(year + 1, 1, 1).Subtract(start.ToDateTime()).TotalDays;
                    }
                    else if(year == end.Year)
                    {
                        daysInLeapYear += end.ToDateTime().Subtract(new System.DateTime(year, 1, 1)).TotalDays;
                    }
                    else
                    {
                        daysInLeapYear += 366d;
                    }
                }
            }
            return daysInLeapYear;
        }
    }
}
