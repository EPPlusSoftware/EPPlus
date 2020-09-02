/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines a pivot table filter type
    /// </summary>
    public enum ePivotTableFilterType
    {
        /// <summary>
        /// A caption filter - Begins With
        /// </summary>
        CaptionBeginsWith,
        /// <summary>
        /// A caption filter - Between
        /// </summary>
        CaptionBetween,
        /// <summary>
        /// A caption filter - Contains
        /// </summary>
        CaptionContains,
        /// <summary>
        /// A caption filter - Ends With
        /// </summary>
        CaptionEndsWith,
        /// <summary>
        /// A caption filter - Equal
        /// </summary>
        CaptionEqual,
        /// <summary>
        /// A caption filter - Greater Than
        /// </summary>
        CaptionGreaterThan,
        /// <summary>
        /// A caption filter - Greater Than Or Equal
        /// </summary>
        CaptionGreaterThanOrEqual,
        /// <summary>
        /// A caption filter - Less Than
        /// </summary>
        CaptionLessThan,
        /// <summary>
        /// A caption filter - Less Than Or Equal
        /// </summary>
        CaptionLessThanOrEqual,
        /// <summary>
        /// A caption filter - Not Begins With
        /// </summary>
        CaptionNotBeginsWith,
        /// <summary>
        /// A caption filter - Not Between
        /// </summary>
        CaptionNotBetween,
        /// <summary>
        /// A caption filter - Not Contains
        /// </summary>
        CaptionNotContains,
        /// <summary>
        /// A caption filter - Not Ends With
        /// </summary>
        CaptionNotEndsWith,
        /// <summary>
        /// A caption filter - Not Equal
        /// </summary>
        CaptionNotEqual,
        /// <summary>
        /// A date filter - Between
        /// </summary>
        DateBetween,
        /// <summary>
        /// A date filter - Equal
        /// </summary>
        DateEqual,
        /// <summary>
        /// A date filter - Newer Than
        /// </summary>
        DateNewerThan,
        /// <summary>
        /// A date filter - Newer Than Or Equal
        /// </summary>
        DateNewerThanOrEqual,
        /// <summary>
        /// A date filter - Not Between
        /// </summary>
        dateNotBetween,
        /// <summary>
        /// A date filter - Not Equal
        /// </summary>
        dateNotEqual,
        /// <summary>
        /// A date filter - Older Than
        /// </summary>
        dateOlderThan,
        /// <summary>
        /// A date filter - Older Than Or Equal
        /// </summary>
        dateOlderThanOrEqual,
        /// <summary>
        /// A date filter - Last Month
        /// </summary>
        lastMonth,
        /// <summary>
        /// A date filter - Last Quarter
        /// </summary>
        lastQuarter,
        /// <summary>
        /// A date filter - Last Week
        /// </summary>
        lastWeek,
        /// <summary>
        /// A date filter - Last Year
        /// </summary>
        lastYear,
        /// <summary>
        /// A date filter - Januari
        /// </summary>
        M1,
        /// <summary>
        /// A date filter - Februari
        /// </summary>
        M2,
        /// <summary>
        /// A date filter - March
        /// </summary>
        M3,
        /// <summary>
        /// A date filter - April
        /// </summary>
        M4,
        /// <summary>
        /// A date filter - May
        /// </summary>
        M5,
        /// <summary>
        /// A date filter - June
        /// </summary>
        M6,
        /// <summary>
        /// A date filter - July
        /// </summary>
        M7,
        /// <summary>
        /// A date filter - August
        /// </summary>
        M8,
        /// <summary>
        /// A date filter - September
        /// </summary>
        M9,
        /// <summary>
        /// A date filter - October
        /// </summary>
        M10, 
        /// <summary>
        /// A date filter - November
        /// </summary>
        M11,
        /// <summary>
        /// A date filter - December
        /// </summary>
        M12,
        /// <summary>
        /// A date filter - Next Month
        /// </summary>
        NextMonth,
        /// <summary>
        /// A date filter - Next Quarter
        /// </summary>
        NextQuarter,
        /// <summary>
        /// A date filter - Next Week
        /// </summary>
        NextWeek,
        /// <summary>
        /// A date filter - Next Year
        /// </summary>
        NextYear,
        /// <summary>
        /// A numeric filter - Percent
        /// </summary>
        Percent,
        /// <summary>
        /// A date filter - The First Quarter
        /// </summary>
        Q1,
        /// <summary>
        /// A date filter - The Second Quarter
        /// </summary>
        Q2,
        /// <summary>
        /// A date filter - The Third Quarter
        /// </summary>
        Q3,
        /// <summary>
        /// A date filter - The Forth Quarter
        /// </summary>
        Q4,
        /// <summary>
        /// A numeric filter - Sum
        /// </summary>
        Sum,
        /// <summary>
        /// A date filter - This Month
        /// </summary>
        ThisMonth,
        /// <summary>
        /// A date filter - This Quarter
        /// </summary>
        ThisQuarter,
        /// <summary>
        /// A date filter - This Week
        /// </summary>
        ThisWeek,
        /// <summary>
        /// A date filter - This Year
        /// </summary>
        ThisYear,
        /// <summary>
        /// A date filter - Today
        /// </summary>
        Today,
        /// <summary>
        /// A date filter - Tomorrow
        /// </summary>
        Tomorrow,
        /// <summary>
        /// Indicates that the filter is unknown
        /// </summary>
        Unknown,
        /// <summary>
        /// A numeric or string filter - Value Between
        /// </summary>
        ValueBetween,
        /// <summary>
        /// A numeric or string filter - Equal
        /// </summary>
        ValueEqual,
        /// <summary>
        /// A numeric or string filter - GreaterThan
        /// </summary>
        ValueGreaterThan,
        /// <summary>
        /// A numeric or string filter - Greater Than Or Equal
        /// </summary>
        ValueGreaterThanOrEqual,
        /// <summary>
        /// A numeric or string filter - Less Than 
        /// </summary>
        ValueLessThan,
        /// <summary>
        /// A numeric or string filter - Less Than Or Equal
        /// </summary>
        ValueLessThanOrEqual,
        /// <summary>
        /// A numeric or string filter - Not Between
        /// </summary>
        ValueNotBetween,
        /// <summary>
        /// A numeric or string filter - Not Equal
        /// </summary>
        ValueNotEqual,
        /// <summary>
        /// A date filter - Year to date
        /// </summary>
        YearToDate,
        /// <summary>
        /// A date filter - Yesterday
        /// </summary>
        Yesterday,
    }
}
