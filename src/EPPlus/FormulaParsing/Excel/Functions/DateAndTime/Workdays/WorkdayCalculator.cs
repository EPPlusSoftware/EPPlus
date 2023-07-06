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
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays
{
    internal class WorkdayCalculator
    {
        private readonly HolidayWeekdays _holidayWeekdays;

        public WorkdayCalculator()
            : this(new HolidayWeekdays())
        {}

        public WorkdayCalculator(HolidayWeekdays holidayWeekdays)
        {
            _holidayWeekdays = holidayWeekdays;
        }

        public WorkdayCalculatorResult CalculateNumberOfWorkdays(DateTime startDate, DateTime endDate)
        {
            var calcDirection = startDate < endDate
                ? WorkdayCalculationDirection.Forward
                : WorkdayCalculationDirection.Backward;
            DateTime calcStartDate;
            DateTime calcEndDate;
            if (calcDirection == WorkdayCalculationDirection.Forward)
            {
                calcStartDate = startDate.Date;
                calcEndDate = endDate.Date;
            }
            else
            {
                calcStartDate = endDate.Date;
                calcEndDate = startDate.Date;
            }
            var nWholeWeeks = (int)calcEndDate.Subtract(calcStartDate).TotalDays/7;
            var workdaysCounted = nWholeWeeks*_holidayWeekdays.NumberOfWorkdaysPerWeek;
            if (!_holidayWeekdays.IsHolidayWeekday(calcStartDate))
            {
                workdaysCounted++;
            }
            var tmpDate = calcStartDate.AddDays(nWholeWeeks*7);
            while (tmpDate < calcEndDate)
            {
                tmpDate = tmpDate.AddDays(1);
                if (!_holidayWeekdays.IsHolidayWeekday(tmpDate))
                {
                    workdaysCounted++;
                }
            }
            return new WorkdayCalculatorResult(workdaysCounted, startDate, endDate, calcDirection);
        }

        public WorkdayCalculatorResult CalculateWorkday(DateTime startDate, int nWorkDays)
        {
            var calcDirection = nWorkDays > 0 ? WorkdayCalculationDirection.Forward : WorkdayCalculationDirection.Backward;
            var direction = (int) calcDirection;
            nWorkDays *= direction;
            var workdaysCounted = 0;
            var tmpDate = startDate;
            
            // calculate whole weeks
            var nWholeWeeks = nWorkDays / _holidayWeekdays.NumberOfWorkdaysPerWeek;
            tmpDate = tmpDate.AddDays(nWholeWeeks * 7 * direction);
            workdaysCounted += nWholeWeeks * _holidayWeekdays.NumberOfWorkdaysPerWeek;

            // calculate the rest
            while (workdaysCounted < nWorkDays)
            {
                tmpDate = tmpDate.AddDays(direction);
                if (!_holidayWeekdays.IsHolidayWeekday(tmpDate)) workdaysCounted++;
            }
            return new WorkdayCalculatorResult(workdaysCounted, startDate, tmpDate, calcDirection);
        }

        public WorkdayCalculatorResult ReduceWorkdaysWithHolidays(WorkdayCalculatorResult calculatedResult,
            FunctionArgument holidayArgument)
        {
            var startDate = calculatedResult.StartDate;
            var endDate = calculatedResult.EndDate;
            var additionalDays = new AdditionalHolidayDays(holidayArgument);
            DateTime calcStartDate;
            DateTime calcEndDate;
            if (startDate < endDate)
            {
                calcStartDate = startDate;
                calcEndDate = endDate;
            }
            else
            {
                calcStartDate = endDate;
                calcEndDate = startDate;
            }
            var nAdditionalHolidayDays = additionalDays.AdditionalDates.Count(x => x >= calcStartDate && x <= calcEndDate && !_holidayWeekdays.IsHolidayWeekday(x));
            return new WorkdayCalculatorResult(calculatedResult.NumberOfWorkdays - nAdditionalHolidayDays, startDate, endDate, calculatedResult.Direction);
        } 

        public WorkdayCalculatorResult AdjustResultWithHolidays(WorkdayCalculatorResult calculatedResult,
                                                         FunctionArgument holidayArgument)
        {
            var startDate = calculatedResult.StartDate;
            var endDate = calculatedResult.EndDate;
            var direction = calculatedResult.Direction;
            var workdaysCounted = calculatedResult.NumberOfWorkdays;
            var additionalDays = new AdditionalHolidayDays(holidayArgument);
            foreach (var date in additionalDays.AdditionalDates)
            {
                if (direction == WorkdayCalculationDirection.Forward && (date < startDate || date > endDate)) continue;
                if (direction == WorkdayCalculationDirection.Backward && (date > startDate || date < endDate)) continue;
                if (_holidayWeekdays.IsHolidayWeekday(date)) continue;
                var tmpDate = _holidayWeekdays.GetNextWorkday(endDate, direction);
                while (additionalDays.AdditionalDates.Contains(tmpDate))
                {
                    tmpDate = _holidayWeekdays.GetNextWorkday(tmpDate, direction);
                }
                workdaysCounted++;
                endDate = tmpDate;
            }

            return new WorkdayCalculatorResult(workdaysCounted, calculatedResult.StartDate, endDate, direction);
        }
    }
}
