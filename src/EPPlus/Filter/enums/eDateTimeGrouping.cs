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
namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Date grouping for a filter
    /// </summary>
    public enum eDateTimeGrouping
    {
        /// <summary>
        /// Group by day
        /// </summary>
        Day,
        /// <summary>
        /// Group by hour
        /// </summary>
        Hour,
        /// <summary>
        /// Group by minute
        /// </summary>
        Minute,
        /// <summary>
        /// Group by month
        /// </summary>
        Month,
        /// <summary>
        /// Group by second
        /// </summary>
        Second,
        /// <summary>
        /// Group by year
        /// </summary>
        Year
    }
}