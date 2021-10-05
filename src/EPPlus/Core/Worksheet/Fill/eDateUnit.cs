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

namespace OfficeOpenXml
{
    /// <summary>
    /// The date units for date fill operations
    /// </summary>
    public enum eDateTimeUnit
    {
        /// <summary>
        /// Adds a Year
        /// </summary>
        Year,
        /// <summary>
        /// Adds a Month
        /// </summary>
        Month,
        /// <summary>
        /// Adds 7 Days
        /// </summary>
        Week,
        /// <summary>
        /// Adds a Day
        /// </summary>
        Day,
        /// <summary>
        /// Adds an Hour
        /// </summary>
        Hour,
        /// <summary>
        /// Adds a Minute
        /// </summary>
        Minute,
        /// <summary>
        /// Adds a Second
        /// </summary>
        Second,
        /// <summary>
        /// Adds ticks
        /// </summary>
        Ticks
    }
}