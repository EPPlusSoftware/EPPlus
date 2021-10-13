/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/11/2021         EPPlus Software AB       EPPlus 5.8
 *************************************************************************************************/

using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Fill
{
    /// <summary>
    /// Parameters for the <see cref="ExcelRangeBase.FillDateTime(Action{FillDateParams})" /> method 
    /// </summary>
    public class FillDateParams : FillParams
    {
        /// <summary>
        /// The start value. If null, the first value in the row/column is used. 
        /// <seealso cref="FillParams.Direction"/>
        /// </summary>
        public DateTime? StartValue { get; set; } = null;
        /// <summary>
        /// When this value is exceeded the fill stops
        /// </summary>
        public DateTime? EndValue { get; set; } = null;
        /// <summary>
        /// The value to add for each step. 
        /// </summary>
        public int StepValue { get; set; } = 1;
        /// <summary>
        /// The date unit added per cell
        /// </summary>
        public eDateTimeUnit DateUnit { get; set; } = eDateTimeUnit.Day;
        /// <summary>
        /// Only fill weekdays
        /// </summary>
        internal HashSet<DayOfWeek> _excludedWeekdays = new HashSet<DayOfWeek>();
        public void SetExcludedWeekdays(params DayOfWeek[] weekdays)
        {
            _excludedWeekdays.UnionWith(weekdays);
        }
        /// <summary>
        /// A list with weekdays treated as holydays.
        /// </summary>
        internal HashSet<DateTime> _holidayCalendar { get; } = new HashSet<DateTime>();
        public void SetHolidayDates(params DateTime[] holidayDates)
        {
            _holidayCalendar.UnionWith(holidayDates);
        }
        public void SetHolidayDates(IEnumerable<DateTime> holidayDates)
        {
            _holidayCalendar.UnionWith(holidayDates);
        }
    }
}
