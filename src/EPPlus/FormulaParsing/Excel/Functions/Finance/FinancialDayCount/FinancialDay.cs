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
    internal abstract class FinancialDay
    {

        public FinancialDay(DateTime date)
        {
            Year = Convert.ToInt16(date.Year);
            Month = Convert.ToInt16(date.Month);
            Day = Convert.ToInt16(date.Day);
        }

        public FinancialDay(int year, int month, int day)
        {
            Year = (short)year;
            Month = (short)month;
            Day = (short)day;
        }

        public override string ToString()
        {
            return $"{Year}-{Month}-{Day}";
        }

        public short Year { get; set; }

        public short Month { get; set; }

        public short Day { get; set; }

        public bool IsLastDayOfFebruary
        {
            get
            {
                return Month == 2 && Day == DateTime.DaysInMonth(Year, Month);
            }
        }

        public bool IsLastDayOfMonth
        {
            get
            {
                return Day == DateTime.DaysInMonth(Year, Month);
            }
        }

        public DateTime ToDateTime()
        {
            return new DateTime(Year, Month, Day);
        }

        public FinancialDay SubtractYears(int years)
        {
            var day = Day;
            if (IsLastDayOfFebruary && DateTime.IsLeapYear(Year) && !DateTime.IsLeapYear(Year + years))
            {
                day -= 1;
            }
            return Factory((short)(Year - years), Month, day);
        }

        public int CompareTo(FinancialDay other)
        {
            if (other is null) return 1;
            if (Year == other.Year && Month == other.Month && Day == other.Day) return 0;
            return ToDateTime().CompareTo(other.ToDateTime());
        }

        public static bool operator >(FinancialDay a, FinancialDay b) => a.CompareTo(b) > 0;

        public static bool operator <(FinancialDay a, FinancialDay b) => a.CompareTo(b) < 0;

        public static bool operator <=(FinancialDay a, FinancialDay b) => a.CompareTo(b) <= 0;

        public static bool operator >=(FinancialDay a, FinancialDay b) => a.CompareTo(b) >= 0;

        public static bool operator ==(FinancialDay a, FinancialDay b)
        {
            if (a is null && b is null) return true;
            if (!(a is null) && b is null) return false;
            if (a is null && !(b is null)) return false;
            return a.CompareTo(b) == 0;
        }

        public static bool operator !=(FinancialDay a, FinancialDay b)
        {
            if (a is null && b is null) return false;
            if (!(a is null) && b is null) return true;
            if (a is null && !(b is null)) return true;
            return a.CompareTo(b) != 0;
        }

        public FinancialDay SubtractMonths(int months, short day)
        {
            var year = Year;
            var actualDay = day;
            var month = Month;
            if (Month - months < 1)
            {
                year -= 1;
                month = Convert.ToInt16(12 - ((Month - months) * -1));
            }
            else
            {
                month = (short)(Month - Convert.ToInt16(months));
            }
            if (IsLastDayOfFebruary && DateTime.IsLeapYear(Year) && !DateTime.IsLeapYear(year))
            {
                actualDay -= 1;
            }
            else if (DateTime.DaysInMonth(year, month) < actualDay)
            {
                actualDay = (short)DateTime.DaysInMonth(year, month);
            }
            return Factory(year, month, actualDay);
        }

        protected abstract FinancialDay Factory(short year, short month, short day);

        internal DayCountBasis GetBasis()
        {
            return Basis;
        }

        protected abstract DayCountBasis Basis { get; }

        /// <summary>
        /// Number of days between two <see cref="FinancialDay"/>s
        /// </summary>
        /// <param name="day">The other day</param>
        /// <returns>Number of days according to the <see cref="DayCountBasis"/> of this day</returns>
        public double SubtractDays(FinancialDay day)
        {
            var financialDays = FinancialDaysFactory.Create(Basis);
            return financialDays.GetDaysBetweenDates(this.ToDateTime(), day.ToDateTime());
        }
        public override bool Equals(object obj)
        {
            if(obj is FinancialDay b)
            {
                return b == this;
            }
            else
            {
                return false;
            }
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
