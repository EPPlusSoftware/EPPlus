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
        CaptionBeginsWith = 0,
        /// <summary>
        /// A caption filter - Between
        /// </summary>
        CaptionBetween = 1,
        /// <summary>
        /// A caption filter - Contains
        /// </summary>
        CaptionContains = 2,
        /// <summary>
        /// A caption filter - Ends With
        /// </summary>
        CaptionEndsWith = 3,
        /// <summary>
        /// A caption filter - Equal
        /// </summary>
        CaptionEqual = 4,
        /// <summary>
        /// A caption filter - Greater Than
        /// </summary>
        CaptionGreaterThan = 5,
        /// <summary>
        /// A caption filter - Greater Than Or Equal
        /// </summary>
        CaptionGreaterThanOrEqual = 6,
        /// <summary>
        /// A caption filter - Less Than
        /// </summary>
        CaptionLessThan = 7,
        /// <summary>
        /// A caption filter - Less Than Or Equal
        /// </summary>
        CaptionLessThanOrEqual = 8,
        /// <summary>
        /// A caption filter - Not Begins With
        /// </summary>
        CaptionNotBeginsWith = 9,
        /// <summary>
        /// A caption filter - Not Between
        /// </summary>
        CaptionNotBetween = 10,
        /// <summary>
        /// A caption filter - Not Contains
        /// </summary>
        CaptionNotContains = 11,
        /// <summary>
        /// A caption filter - Not Ends With
        /// </summary>
        CaptionNotEndsWith = 12,
        /// <summary>
        /// A caption filter - Not Equal
        /// </summary>
        CaptionNotEqual = 13,
        /// <summary>
        /// A date filter - Between
        /// </summary>
        DateBetween = 100,
        /// <summary>
        /// A date filter - Equal
        /// </summary>
        DateEqual = 101,
        /// <summary>
        /// A date filter - Newer Than
        /// </summary>
        DateNewerThan = 102,
        /// <summary>
        /// A date filter - Newer Than Or Equal
        /// </summary>
        DateNewerThanOrEqual = 103,
        /// <summary>
        /// A date filter - Not Between
        /// </summary>
        DateNotBetween = 104,
        /// <summary>
        /// A date filter - Not Equal
        /// </summary>
        DateNotEqual = 105,
        /// <summary>
        /// A date filter - Older Than
        /// </summary>
        DateOlderThan = 106,
        /// <summary>
        /// A date filter - Older Than Or Equal
        /// </summary>
        DateOlderThanOrEqual = 107,
        /// <summary>
        /// A date filter - Last Month
        /// </summary>
        LastMonth = 108,
        /// <summary>
        /// A date filter - Last Quarter
        /// </summary>
        LastQuarter = 200,
        /// <summary>
        /// A date filter - Last Week
        /// </summary>
        LastWeek = 201,
        /// <summary>
        /// A date filter - Last Year
        /// </summary>
        LastYear = 202,
        /// <summary>
        /// A date filter - Januari
        /// </summary>
        M1 = 203,
        /// <summary>
        /// A date filter - Februari
        /// </summary>
        M2 = 204,
        /// <summary>
        /// A date filter - March
        /// </summary>
        M3 = 205,
        /// <summary>
        /// A date filter - April
        /// </summary>
        M4 = 206,
        /// <summary>
        /// A date filter - May
        /// </summary>
        M5 = 207,
        /// <summary>
        /// A date filter - June
        /// </summary>
        M6 = 208,
        /// <summary>
        /// A date filter - July
        /// </summary>
        M7 = 209,
        /// <summary>
        /// A date filter - August
        /// </summary>
        M8 = 210,
        /// <summary>
        /// A date filter - September
        /// </summary>
        M9 = 211,
        /// <summary>
        /// A date filter - October
        /// </summary>
        M10 = 212,
        /// <summary>
        /// A date filter - November
        /// </summary>
        M11 = 213,
        /// <summary>
        /// A date filter - December
        /// </summary>
        M12 = 214,
        /// <summary>
        /// A date filter - Next Month
        /// </summary>
        NextMonth = 215,
        /// <summary>
        /// A date filter - Next Quarter
        /// </summary>
        NextQuarter = 216,
        /// <summary>
        /// A date filter - Next Week
        /// </summary>
        NextWeek = 217,
        /// <summary>
        /// A date filter - Next Year
        /// </summary>
        NextYear = 218,
        /// <summary>
        /// A date filter - The First Quarter
        /// </summary>
        Q1 = 219,
        /// <summary>
        /// A date filter - The Second Quarter
        /// </summary>
        Q2 = 220,
        /// <summary>
        /// A date filter - The Third Quarter
        /// </summary>
        Q3 = 221,
        /// <summary>
        /// A date filter - The Forth Quarter
        /// </summary>
        Q4 = 222,
        /// <summary>
        /// A date filter - This Month
        /// </summary>
        ThisMonth = 223,
        /// <summary>
        /// A date filter - This Quarter
        /// </summary>
        ThisQuarter = 224,
        /// <summary>
        /// A date filter - This Week
        /// </summary>
        ThisWeek = 225,
        /// <summary>
        /// A date filter - This Year
        /// </summary>
        ThisYear = 226,
        /// <summary>
        /// A date filter - Today
        /// </summary>
        Today = 227,
        /// <summary>
        /// A date filter - Tomorrow
        /// </summary>
        Tomorrow = 228,
        /// <summary>
        /// A date filter - Year to date
        /// </summary>
        YearToDate = 229,
        /// <summary>
        /// A date filter - Yesterday
        /// </summary>
        Yesterday = 230,
        /// <summary>
        /// Indicates that the filter is unknown
        /// </summary>
        Unknown = -1,
        /// <summary>
        /// A numeric filter - Sum
        /// </summary>
        Sum = 300,
        /// <summary>
        /// A numeric filter - Percent
        /// </summary>
        Percent = 301,
        /// <summary>
        /// A numeric or string filter - Value Between
        /// </summary>
        ValueBetween = 302,
        /// <summary>
        /// A numeric or string filter - Equal
        /// </summary>
        ValueEqual = 303,
        /// <summary>
        /// A numeric or string filter - GreaterThan
        /// </summary>
        ValueGreaterThan = 304,
        /// <summary>
        /// A numeric or string filter - Greater Than Or Equal
        /// </summary>
        ValueGreaterThanOrEqual = 305,
        /// <summary>
        /// A numeric or string filter - Less Than 
        /// </summary>
        ValueLessThan = 306,
        /// <summary>
        /// A numeric or string filter - Less Than Or Equal
        /// </summary>
        ValueLessThanOrEqual = 307,
        /// <summary>
        /// A numeric or string filter - Not Between
        /// </summary>
        ValueNotBetween = 308,
        /// <summary>
        /// A numeric or string filter - Not Equal
        /// </summary>
        ValueNotEqual = 308
    }
}
