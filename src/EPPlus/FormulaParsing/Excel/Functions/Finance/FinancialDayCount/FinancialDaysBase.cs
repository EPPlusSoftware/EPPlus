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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
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

        public IEnumerable<FinancialPeriod> GetCouponPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency)
        {
            var periods = new List<FinancialPeriod>();
            var tmpDay = settlement;
            while(tmpDay > date)
            {
                var periodEndDay = tmpDay;
                switch (frequency)
                {
                    case 1:
                        tmpDay = tmpDay.SubtractYears(1);
                        break;
                    case 2:
                        tmpDay = tmpDay.SubtractMonths(6, settlement.Day);
                        break;
                    case 4:
                        tmpDay = tmpDay.SubtractMonths(3, settlement.Day);
                        break;
                    default:
                        throw new ArgumentException("frequency");
                }
                periods.Add(new FinancialPeriod(tmpDay, periodEndDay));
            }
            return periods;
        }

        private FinancialPeriod CreateCalendarPeriod(System.DateTime startDate, int frequency, DayCountBasis basis, bool createFuturePeriod)
        {
            var d1 = System.DateTime.MinValue;
            var factor = createFuturePeriod ? 1 : -1;
            switch(frequency)
            {
                case 1:
                    d1 = startDate.AddYears(1 * factor);
                    break;
                case 2:
                    d1 = startDate.AddMonths(6 * factor);
                    break;
                case 4:
                    d1 = startDate.AddMonths(3 * factor);
                    break;
                default:
                    throw new ArgumentException("frequency");
            }
            if(createFuturePeriod)
            {
                return FinancialDayFactory.CreatePeriod(startDate, d1, basis);
            }
            else
            {
                return FinancialDayFactory.CreatePeriod(d1, startDate, basis);
            }
        }

        private FinancialPeriod GetSettlementCalendarYearPeriod(FinancialDay date, int frequency)
        {
            System.DateTime startDate = default(System.DateTime);
            if (frequency == 1)
            {
                startDate = new System.DateTime(date.Year, 1, 1);
            }
            else if(frequency == 2)
            {
                if(date.Month < 7)
                {
                    startDate = new System.DateTime(date.Year, 1, 1);
                }
                else
                {
                    startDate = new System.DateTime(date.Year, 7, 1);
                }
            }
            else if(frequency == 4)
            {
                if (date.Month > 9)
                {
                    startDate = new System.DateTime(date.Year, 10, 1);
                }
                else if(date.Month > 6)
                {
                    startDate = new System.DateTime(date.Year, 7, 1);
                }
                else if(date.Month > 3)
                {
                    startDate = new System.DateTime(date.Year, 4, 1);
                }
                else
                {
                    startDate = new System.DateTime(date.Year, 1, 1);
                }
            }
            else
            {
                throw new ArgumentException("frequency");
            }
            return CreateCalendarPeriod(startDate, frequency, date.GetBasis(), true);
        }

        public IEnumerable<FinancialPeriod> GetCalendarYearPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency)
        {
            return GetCalendarYearPeriodsBackwards(settlement, date, frequency, 0);
        }

        public IEnumerable<FinancialPeriod> GetCalendarYearPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency, int additionalPeriods)
        {
            var periods = new List<FinancialPeriod>();
            var period = GetSettlementCalendarYearPeriod(settlement, frequency);
            periods.Add(period);
            while (period.Start > date)
            {
                var dt = period.Start.ToDateTime();
                period = CreateCalendarPeriod(dt, frequency, date.GetBasis(), false);
                periods.Add(period);
            }
            for(var x = 0; x < additionalPeriods; x++)
            {
                var tmpDate = periods.Last().Start.ToDateTime();
                var p = CreateCalendarPeriod(tmpDate, frequency, date.GetBasis(), false);
                periods.Add(p);
            }
            return periods;
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

        protected virtual double GetDaysBetweenDates(FinancialDay start, FinancialDay end, int basis, bool returnZeroIfNegative)
        {
            var result = (basis * (end.Year - start.Year) + 30 * (end.Month - start.Month) + ((end.Day > 30 ? 30 : end.Day) - (start.Day > 30 ? 30 : start.Day)));
            if (returnZeroIfNegative && result < 0 )
            {
                return 0d;
            }
            else
            {
                return result;
            }
        }

        protected virtual double GetDaysBetweenDates(FinancialDay start, FinancialDay end, int basis)
        {
            return GetDaysBetweenDates(start, end, basis, false);
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
