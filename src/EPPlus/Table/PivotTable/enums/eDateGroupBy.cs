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

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Data grouping
    /// </summary>
    [Flags]
    public enum eDateGroupBy
    {
        /// <summary>
        /// Group by years
        /// </summary>
        Years = 1,
        /// <summary>
        /// Group by  quarters
        /// </summary>
        Quarters = 2,
        /// <summary>
        /// Group by months
        /// </summary>
        Months = 4,
        /// <summary>
        /// Group by days
        /// </summary>
        Days = 8,
        /// <summary>
        /// Group by hours
        /// </summary>
        Hours = 16,
        /// <summary>
        /// Group by minutes
        /// </summary>
        Minutes = 32,
        /// <summary>
        /// Group by seconds
        /// </summary>
        Seconds = 64
    }
}