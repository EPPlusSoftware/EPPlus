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
    public class HolidayWeekdays
    {
        private readonly List<DayOfWeek> _holidayDays = new List<DayOfWeek>();

        public HolidayWeekdays()
            :this(DayOfWeek.Saturday, DayOfWeek.Sunday)
        {
            
        }

        public int NumberOfWorkdaysPerWeek => 7 - _holidayDays.Count;

        public HolidayWeekdays(params DayOfWeek[] holidayDays)
        {
            foreach (var dayOfWeek in holidayDays)
            {
                _holidayDays.Add(dayOfWeek);
            }
        }

        public bool IsHolidayWeekday(DateTime dateTime)
        {
            return _holidayDays.Contains(dateTime.DayOfWeek);
        }

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
