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
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays
{
    /// <summary>
    /// Workday calculator result
    /// </summary>
    internal class WorkdayCalculatorResult
    {
        /// <summary>
        /// Constructor. Calculate workdays
        /// </summary>
        /// <param name="numberOfWorkdays"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="direction"></param>
        public WorkdayCalculatorResult(int numberOfWorkdays, DateTime startDate, DateTime endDate, WorkdayCalculationDirection direction)
        {
            NumberOfWorkdays = numberOfWorkdays;
            StartDate = startDate;
            EndDate = endDate;
            Direction = direction;
        }

        /// <summary>
        /// Number of Workdays
        /// </summary>
        public int NumberOfWorkdays { get; }

        /// <summary>
        /// Start date
        /// </summary>
        public DateTime StartDate { get; }

        /// <summary>
        /// End date
        /// </summary>
        public DateTime EndDate { get; }

        /// <summary>
        /// Direction to look for workdays in
        /// </summary>
        public WorkdayCalculationDirection Direction { get; set; }
    }
}
