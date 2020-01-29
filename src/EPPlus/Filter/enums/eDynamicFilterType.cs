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
    /// Dynamic filter types.
    /// A dynamic filter returns a result set which might vary due to a change in the data itself.
    /// </summary>
    public enum eDynamicFilterType
    {
        /// <summary>
        /// Shows values that are above average. 
        /// </summary>
        AboveAverage,
        /// <summary>
        /// Shows values that are below average. 
        /// </summary>
        BelowAverage,
        /// <summary>
        /// Shows last month's dates.
        /// </summary>
        LastMonth,
        /// <summary>
        /// Shows last calendar quarter's dates.
        /// </summary>
        LastQuarter,
        /// <summary>
        /// Shows last week's dates, using Sunday as the first weekday.
        /// </summary>
        LastWeek,
        /// <summary>
        ///  Shows last year's dates.
        /// </summary>
        LastYear,
        /// <summary>
        /// Shows the dates that are in January, regardless of year.
        /// </summary>
        M1,
        /// <summary>
        /// Shows the dates that are in February, regardless of year. 
        /// </summary>
        M2,
        /// <summary>
        /// Shows the dates that are in March, regardless of year.
        /// </summary>
        M3, 
        /// <summary>
        /// Shows the dates that are in April, regardless of year.
        /// </summary>
        M4, 
        /// <summary>
        /// Shows the dates that are in May, regardless of year.
        /// </summary>
        M5, 
        /// <summary>
        /// Shows the dates that are in June, regardless of year.
        /// </summary>
        M6, 
        /// <summary>
        /// Shows the dates that are in July, regardless of year.
        /// </summary>
        M7, 
        /// <summary>
        /// Shows the dates that are in August, regardless of year.
        /// </summary>
        M8, 
        /// <summary>
        /// Shows the dates that are in September, regardless of
        /// </summary>
        M9, 
        /// <summary>
        /// Shows the dates that are in October, regardless of year.
        /// </summary>
        M10, 
        /// <summary>
        /// Shows the dates that are in November, regardless of year.
        /// </summary>
        M11, 
        /// <summary>
        /// Shows the dates that are in December, regardless of year.
        /// </summary>
        M12,
        /// <summary>
        /// Shows next month's dates.
        /// </summary>
        NextMonth, 
        /// <summary>
        /// Shows next calendar quarter's dates.
        /// </summary>
        NextQuarter,
        /// <summary>
        /// Shows next week's dates, using Sunday as the firstweekday.
        /// </summary>
        NextWeek, 
        /// <summary>
        /// Shows next year's dates.
        /// </summary>
        NextYear, 
        /// <summary>
        /// No filter
        /// </summary>
        Null, 
        /// <summary>
        /// Shows the dates that are in the 1st calendar quarter, regardless of year. 
        /// </summary>
        Q1,
        /// <summary>
        /// Shows the dates that are in the 2nd calendar quarter, regardless of year. 
        /// </summary>
        Q2,
        /// <summary>
        /// Shows the dates that are in the 3rd calendar quarter, regardless of year. 
        /// </summary>
        Q3,
        /// <summary>
        /// Shows the dates that are in the 4th calendar quarter, regardless of year.
        /// </summary>
        Q4,
        /// <summary>
        /// Shows this month's dates.
        /// </summary>
        ThisMonth,
        /// <summary>
        /// Shows this calendar quarter's dates.
        /// </summary>
        ThisQuarter,
        /// <summary>
        /// Shows this week's dates, using Sunday as the first weekday.
        /// </summary>
        ThisWeek,
        /// <summary>
        /// Shows this year's dates.
        /// </summary>
        ThisYear,
        /// <summary>
        /// Shows today's dates.
        /// </summary>
        Today,
        /// <summary>
        /// Shows tomorrow's dates.
        /// </summary>
        Tomorrow,
        /// <summary>
        /// Shows the dates between the beginning of the year and today, inclusive.
        /// </summary>
        YearToDate,
        /// <summary>
        /// Shows yesterday's dates.
        /// </summary>
        Yesterday
    }
}