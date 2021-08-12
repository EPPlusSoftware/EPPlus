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

namespace OfficeOpenXml
{
    public class FillDateParams
    {
        /// <summary>
        /// The start value. If null, the first value in the row/column is used. 
        /// <seealso cref="Direction"/>
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
        /// The direction of the fill
        /// </summary>
        public eFillDirection Direction { get; set; } = eFillDirection.Column;
        /// <summary>
        /// The date unit added per cell
        /// </summary>
        public eDateUnit DateUnit { get; set; } = eDateUnit.Day;
        /// <summary>
        /// Only fill weekdays
        /// </summary>
        public bool WeekdaysOnly { get; set; } = false;
        /// <summary>
        /// A list with weekdays treated as holydays.
        /// </summary>
        public HashSet<DateTime> HolidayCalendar { get; } = new HashSet<DateTime>();
    }
}
