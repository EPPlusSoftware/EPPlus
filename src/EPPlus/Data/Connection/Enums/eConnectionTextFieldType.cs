/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Specifies the external connection data type and its date format order,
    /// used when importing or interpreting external data fields.
    /// </summary>
    public enum eConnectionTextFieldType
    {
        /// <summary>
        /// Decides the best-fit data type based on the content.
        /// </summary>
        General,
        /// <summary>
        /// Skip this field entirely — do not import it.
        /// </summary>
        SkipField,

        /// <summary>
        /// Field contains plain text.
        /// </summary>
        Text,
        /// <summary>
        /// Field contains a date in the order: day, month, year.
        /// </summary>
        DayMonthYear,

        /// <summary>
        /// Field contains a date in the order: day, year, month.
        /// </summary>
        DayYearMonth,

        /// <summary>
        /// Field contains an East Asian date in the order: year, month, day.
        /// </summary>
        EastAsianYearMonthDay,

        /// <summary>
        /// Field contains a date in the order: month, day, year.
        /// </summary>
        MonthDayYear,

        /// <summary>
        /// Field contains a date in the order: month, year, day.
        /// </summary>
        MonthYearDay,

        /// <summary>
        /// Field contains a date in the order: year, day, month.
        /// </summary>
        YearDayMonth,

        /// <summary>
        /// Field contains a date in the order: year, month, day.
        /// </summary>
        YearMonthDay
    }
}