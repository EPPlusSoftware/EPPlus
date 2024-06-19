/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/12/2023         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal class DateTimeUtil
    {
        internal static void GetQuarterDates(DateTime date, out DateTime startDate, out DateTime endDate)
        {
            var quarter = ((date.Month - (date.Month - 1) % 3) + 1) / 3;
            startDate = new DateTime(date.Year, (quarter * 3) + 1, 1);
            endDate = startDate.AddMonths(3).AddDays(quarter - 1);
        }

        internal static void GetWeekDates(DateTime date, out DateTime startDate, out DateTime endDate)
        {
            while (date.DayOfWeek != DayOfWeek.Sunday)
            {
                date = date.AddDays(-1);
            }
            startDate = date;
            endDate = date.AddDays(6);
        }
    }
}
