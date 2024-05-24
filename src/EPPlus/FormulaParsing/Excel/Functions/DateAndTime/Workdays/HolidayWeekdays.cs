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
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays
{
    /// <summary>
    /// Holiday weekdays for datetime functions
    /// </summary>
    internal class HolidayWeekdays
    {
        private readonly List<DayOfWeek> _holidayDays = new List<DayOfWeek>();

        /// <summary>
        /// Empty constuctor. Sets saturday and sunday to holiday days automatically.
        /// </summary>
        public HolidayWeekdays()
            :this(DayOfWeek.Saturday, DayOfWeek.Sunday)
        {
            
        }

        /// <summary>
        /// Defined as 7 - number of holidayDays
        /// </summary>
        public int NumberOfWorkdaysPerWeek => 7 - _holidayDays.Count;

        /// <summary>
        /// Define holiday weekdays by input array
        /// </summary>
        /// <param name="holidayDays"></param>
        public HolidayWeekdays(params DayOfWeek[] holidayDays)
        {
            foreach (var dayOfWeek in holidayDays)
            {
                _holidayDays.Add(dayOfWeek);
            }
        }
        /// <summary>
        /// Check wether given datetime is a holidayWeekday or not
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public bool IsHolidayWeekday(DateTime dateTime)
        {
            return _holidayDays.Contains(dateTime.DayOfWeek);
        }
        /// <summary>
        /// Adjust result with holidays.
        /// </summary>
        /// <param name="resultDate"></param>
        /// <param name="arguments"></param>
        /// <returns></returns>
        public DateTime AdjustResultWithHolidays(DateTime resultDate,
                                                         IEnumerable<FunctionArgument> arguments)
        {
            if (arguments.Count() == 2) return resultDate;
            var holidays = arguments.ElementAt(2).Value as IEnumerable<FunctionArgument>;
            if (holidays != null)
            {
                foreach (var arg in holidays)
                {
                    if (ConvertUtil.IsNumericOrDate(arg.Value))
                    {
                        var dateSerial = ConvertUtil.GetValueDouble(arg.Value);
                        var holidayDate = DateTime.FromOADate(dateSerial);
                        if (!IsHolidayWeekday(holidayDate))
                        {
                            resultDate = resultDate.AddDays(1);
                        }
                    }
                }
            }
            else
            {
                var range = arguments.ElementAt(2).Value as IRangeInfo;
                if (range != null)
                {
                    foreach (var cell in range)
                    {
                        if (ConvertUtil.IsNumericOrDate(cell.Value))
                        {
                            var dateSerial = ConvertUtil.GetValueDouble(cell.Value);
                            var holidayDate = DateTime.FromOADate(dateSerial);
                            if (!IsHolidayWeekday(holidayDate))
                            {
                                resultDate = resultDate.AddDays(1);
                            }
                        }
                    }
                }
            }
            return resultDate;
        }
        /// <summary>
        /// Gets the next datetime workday
        /// </summary>
        /// <param name="date"></param>
        /// <param name="direction"></param>
        /// <returns></returns>
        public DateTime GetNextWorkday(DateTime date, WorkdayCalculationDirection direction = WorkdayCalculationDirection.Forward)
        {
            var changeParam = (int)direction;
            var tmpDate = date.AddDays(changeParam);
            while (IsHolidayWeekday(tmpDate))
            {
                tmpDate = tmpDate.AddDays(changeParam);
            }
            return tmpDate;
        }
    }
}
