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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays
{
    /// <summary>
    /// Factory class for holidayWeekdays
    /// </summary>
    internal class HolidayWeekdaysFactory
    {
        private readonly DayOfWeek[] _dayOfWeekArray = new DayOfWeek[]
        {
            DayOfWeek.Monday, 
            DayOfWeek.Tuesday, 
            DayOfWeek.Wednesday, 
            DayOfWeek.Thursday,
            DayOfWeek.Friday, 
            DayOfWeek.Saturday,
            DayOfWeek.Sunday
        };

        /// <summary>
        /// Create from string
        /// </summary>
        /// <param name="weekdays"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public HolidayWeekdays Create(string weekdays)
        {
            if(string.IsNullOrEmpty(weekdays) || weekdays.Length != 7)
                throw new ArgumentException("Illegal weekday string", nameof(Weekday));

            var retVal = new List<DayOfWeek>();
            var arr = weekdays.ToCharArray();
            for(var i = 0; i < arr.Length;i++)
            {
                var ch = arr[i];
                if (ch == '1')
                {
                    retVal.Add(_dayOfWeekArray[i]);
                }
            }
            return new HolidayWeekdays(retVal.ToArray());
        }
        /// <summary>
        /// Create from code
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public HolidayWeekdays Create(int code)
        {
            switch (code)
            {
                case 1:
                    return new HolidayWeekdays(DayOfWeek.Saturday, DayOfWeek.Sunday);
                case 2:
                    return new HolidayWeekdays(DayOfWeek.Sunday, DayOfWeek.Monday);
                case 3:
                    return new HolidayWeekdays(DayOfWeek.Monday, DayOfWeek.Tuesday);
                case 4:
                    return new HolidayWeekdays(DayOfWeek.Tuesday, DayOfWeek.Wednesday);
                case 5:
                    return new HolidayWeekdays(DayOfWeek.Wednesday, DayOfWeek.Thursday);
                case 6:
                    return new HolidayWeekdays(DayOfWeek.Thursday, DayOfWeek.Friday);
                case 7:
                    return new HolidayWeekdays(DayOfWeek.Friday, DayOfWeek.Saturday);
                case 11:
                    return new HolidayWeekdays(DayOfWeek.Sunday);
                case 12:
                    return new HolidayWeekdays(DayOfWeek.Monday);
                case 13:
                    return new HolidayWeekdays(DayOfWeek.Tuesday);
                case 14:
                    return new HolidayWeekdays(DayOfWeek.Wednesday);
                case 15:
                    return new HolidayWeekdays(DayOfWeek.Thursday);
                case 16:
                    return new HolidayWeekdays(DayOfWeek.Friday);
                case 17:
                    return new HolidayWeekdays(DayOfWeek.Saturday);
                default:
                    throw new ArgumentException("Invalid code supplied to HolidayWeekdaysFactory: " + code);
            }
        }
    }
}
