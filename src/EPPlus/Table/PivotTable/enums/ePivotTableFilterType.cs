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
        LastMonth = 200,
        /// <summary>
        /// A date filter - Last Quarter
        /// </summary>
        LastQuarter = 201,
        /// <summary>
        /// A date filter - Last Week
        /// </summary>
        LastWeek = 202,
        /// <summary>
        /// A date filter - Last Year
        /// </summary>
        LastYear = 203,
        /// <summary>
        /// A date filter - Januari
        /// </summary>
        M1 = 204,
        /// <summary>
        /// A date filter - Februari
        /// </summary>
        M2 = 205,
        /// <summary>
        /// A date filter - March
        /// </summary>
        M3 = 206,
        /// <summary>
        /// A date filter - April
        /// </summary>
        M4 = 207,
        /// <summary>
        /// A date filter - May
        /// </summary>
        M5 = 208,
        /// <summary>
        /// A date filter - June
        /// </summary>
        M6 = 209,
        /// <summary>
        /// A date filter - July
        /// </summary>
        M7 = 210,
        /// <summary>
        /// A date filter - August
        /// </summary>
        M8 = 211,
        /// <summary>
        /// A date filter - September
        /// </summary>
        M9 = 212,
        /// <summary>
        /// A date filter - October
        /// </summary>
        M10 = 213,
        /// <summary>
        /// A date filter - November
        /// </summary>
        M11 = 214,
        /// <summary>
        /// A date filter - December
        /// </summary>
        M12 = 215,
        /// <summary>
        /// A date filter - Next Month
        /// </summary>
        NextMonth = 216,
        /// <summary>
        /// A date filter - Next Quarter
        /// </summary>
        NextQuarter = 217,
        /// <summary>
        /// A date filter - Next Week
        /// </summary>
        NextWeek = 218,
        /// <summary>
        /// A date filter - Next Year
        /// </summary>
        NextYear = 219,
        /// <summary>
        /// A date filter - The First Quarter
        /// </summary>
        Q1 = 220,
        /// <summary>
        /// A date filter - The Second Quarter
        /// </summary>
        Q2 = 221,
        /// <summary>
        /// A date filter - The Third Quarter
        /// </summary>
        Q3 = 222,
        /// <summary>
        /// A date filter - The Forth Quarter
        /// </summary>
        Q4 = 223,
        /// <summary>
        /// A date filter - This Month
        /// </summary>
        ThisMonth = 224,
        /// <summary>
        /// A date filter - This Quarter
        /// </summary>
        ThisQuarter = 225,
        /// <summary>
        /// A date filter - This Week
        /// </summary>
        ThisWeek = 226,
        /// <summary>
        /// A date filter - This Year
        /// </summary>
        ThisYear = 227,
        /// <summary>
        /// A date filter - Today
        /// </summary>
        Today = 228,
        /// <summary>
        /// A date filter - Tomorrow
        /// </summary>
        Tomorrow = 229,
        /// <summary>
        /// A date filter - Year to date
        /// </summary>
        YearToDate = 230,
        /// <summary>
        /// A date filter - Yesterday
        /// </summary>
        Yesterday = 231,
        /// <summary>
        /// Indicates that the filter is unknown
        /// </summary>
        Unknown = -1,
        /// <summary>
        /// A numeric or string filter - Value Between
        /// </summary>
        ValueBetween = 300,
        /// <summary>
        /// A numeric or string filter - Equal
        /// </summary>
        ValueEqual = 301,
        /// <summary>
        /// A numeric or string filter - GreaterThan
        /// </summary>
        ValueGreaterThan = 302,
        /// <summary>
        /// A numeric or string filter - Greater Than Or Equal
        /// </summary>
        ValueGreaterThanOrEqual = 303,
        /// <summary>
        /// A numeric or string filter - Less Than 
        /// </summary>
        ValueLessThan = 304,
        /// <summary>
        /// A numeric or string filter - Less Than Or Equal
        /// </summary>
        ValueLessThanOrEqual = 305,
        /// <summary>
        /// A numeric or string filter - Not Between
        /// </summary>
        ValueNotBetween = 306,
        /// <summary>
        /// A numeric or string filter - Not Equal
        /// </summary>
        ValueNotEqual = 307,
        /// <summary>
        /// A top/bottom filter - Count
        /// </summary>
        Count = 400,
        /// <summary>
        /// A top/bottom filter - Sum
        /// </summary>
        Sum = 401,
        /// <summary>
        /// A top/bottom filter - Percent
        /// </summary>
        Percent = 402,
    }
    public enum ePivotTableCaptionFilterType
    {
        /// <summary>
        /// A caption filter - Begins With
        /// </summary>
        CaptionBeginsWith = ePivotTableFilterType.CaptionBeginsWith,
        /// <summary>
        /// A caption filter - Between
        /// </summary>
        CaptionBetween = ePivotTableFilterType.CaptionBetween,
        /// <summary>
        /// A caption filter - Contains
        /// </summary>
        CaptionContains = ePivotTableFilterType.CaptionContains,
        /// <summary>
        /// A caption filter - Ends With
        /// </summary>
        CaptionEndsWith = ePivotTableFilterType.CaptionEndsWith,
        /// <summary>
        /// A caption filter - Equal
        /// </summary>
        CaptionEqual = ePivotTableFilterType.CaptionEqual,
        /// <summary>
        /// A caption filter - Greater Than
        /// </summary>
        CaptionGreaterThan = ePivotTableFilterType.CaptionGreaterThan,
        /// <summary>
        /// A caption filter - Greater Than Or Equal
        /// </summary>
        CaptionGreaterThanOrEqual = ePivotTableFilterType.CaptionGreaterThanOrEqual,
        /// <summary>
        /// A caption filter - Less Than
        /// </summary>
        CaptionLessThan = ePivotTableFilterType.CaptionLessThan,
        /// <summary>
        /// A caption filter - Less Than Or Equal
        /// </summary>
        CaptionLessThanOrEqual = ePivotTableFilterType.CaptionLessThanOrEqual,
        /// <summary>
        /// A caption filter - Not Begins With
        /// </summary>
        CaptionNotBeginsWith = ePivotTableFilterType.CaptionNotBeginsWith,
        /// <summary>
        /// A caption filter - Not Between
        /// </summary>
        CaptionNotBetween = ePivotTableFilterType.CaptionNotBetween,
        /// <summary>
        /// A caption filter - Not Contains
        /// </summary>
        CaptionNotContains = ePivotTableFilterType.CaptionNotContains,
        /// <summary>
        /// A caption filter - Not Ends With
        /// </summary>
        CaptionNotEndsWith = ePivotTableFilterType.CaptionNotEndsWith,
        /// <summary>
        /// A caption filter - Not Equal
        /// </summary>
        CaptionNotEqual = ePivotTableFilterType.CaptionNotEqual,
    }
    public enum ePivotTableDateValueFilterType
    {
        /// <summary>
        /// A date filter - Between
        /// </summary>
        DateBetween = ePivotTableFilterType.DateBetween,
        /// <summary>
        /// A date filter - Equal
        /// </summary>
        DateEqual = ePivotTableFilterType.DateEqual,
        /// <summary>
        /// A date filter - Newer Than
        /// </summary>
        DateNewerThan = ePivotTableFilterType.DateNewerThan,
        /// <summary>
        /// A date filter - Newer Than Or Equal
        /// </summary>
        DateNewerThanOrEqual = ePivotTableFilterType.DateNewerThanOrEqual,
        /// <summary>
        /// A date filter - Not Between
        /// </summary>
        DateNotBetween = ePivotTableFilterType.DateNotBetween,
        /// <summary>
        /// A date filter - Not Equal
        /// </summary>
        DateNotEqual = ePivotTableFilterType.DateNotEqual,
        /// <summary>
        /// A date filter - Older Than
        /// </summary>
        DateOlderThan = ePivotTableFilterType.DateOlderThan,
        /// <summary>
        /// A date filter - Older Than Or Equal
        /// </summary>
        DateOlderThanOrEqual = ePivotTableFilterType.DateOlderThanOrEqual,
    }
    public enum ePivotTableDatePeriodFilterType
    {
        /// <summary>
        /// A date filter - Last Month
        /// </summary>
        LastMonth = ePivotTableFilterType.LastMonth,
        /// <summary>
        /// A date filter - Last Quarter
        /// </summary>
        LastQuarter = ePivotTableFilterType.LastQuarter,
        /// <summary>
        /// A date filter - Last Week
        /// </summary>
        LastWeek = ePivotTableFilterType.LastWeek,
        /// <summary>
        /// A date filter - Last Year
        /// </summary>
        LastYear = ePivotTableFilterType.LastYear,
        /// <summary>
        /// A date filter - Januari
        /// </summary>
        M1 = ePivotTableFilterType.M1,
        /// <summary>
        /// A date filter - Februari
        /// </summary>
        M2 = ePivotTableFilterType.M2,
        /// <summary>
        /// A date filter - March
        /// </summary>
        M3 = ePivotTableFilterType.M3,
        /// <summary>
        /// A date filter - April
        /// </summary>
        M4 = ePivotTableFilterType.M4,
        /// <summary>
        /// A date filter - May
        /// </summary>
        M5 = ePivotTableFilterType.M5,
        /// <summary>
        /// A date filter - June
        /// </summary>
        M6 = ePivotTableFilterType.M6,
        /// <summary>
        /// A date filter - July
        /// </summary>
        M7 = ePivotTableFilterType.M7,
        /// <summary>
        /// A date filter - August
        /// </summary>
        M8 = ePivotTableFilterType.M8,
        /// <summary>
        /// A date filter - September
        /// </summary>
        M9 = ePivotTableFilterType.M9,
        /// <summary>
        /// A date filter - October
        /// </summary>
        M10 = ePivotTableFilterType.M10,
        /// <summary>
        /// A date filter - November
        /// </summary>
        M11 = ePivotTableFilterType.M11,
        /// <summary>
        /// A date filter - December
        /// </summary>
        M12 = ePivotTableFilterType.M12,
        /// <summary>
        /// A date filter - Next Month
        /// </summary>
        NextMonth = ePivotTableFilterType.NextMonth,
        /// <summary>
        /// A date filter - Next Quarter
        /// </summary>
        NextQuarter = ePivotTableFilterType.NextQuarter,
        /// <summary>
        /// A date filter - Next Week
        /// </summary>
        NextWeek = ePivotTableFilterType.NextWeek,
        /// <summary>
        /// A date filter - Next Year
        /// </summary>
        NextYear = ePivotTableFilterType.NextYear,
        /// <summary>
        /// A date filter - The First Quarter
        /// </summary>
        Q1 = ePivotTableFilterType.Q1,
        /// <summary>
        /// A date filter - The Second Quarter
        /// </summary>
        Q2 = ePivotTableFilterType.Q2,
        /// <summary>
        /// A date filter - The Third Quarter
        /// </summary>
        Q3 = ePivotTableFilterType.Q3,
        /// <summary>
        /// A date filter - The Forth Quarter
        /// </summary>
        Q4 = ePivotTableFilterType.Q4,
        /// <summary>
        /// A date filter - This Month
        /// </summary>
        ThisMonth = ePivotTableFilterType.ThisMonth,
        /// <summary>
        /// A date filter - This Quarter
        /// </summary>
        ThisQuarter = ePivotTableFilterType.ThisQuarter,
        /// <summary>
        /// A date filter - This Week
        /// </summary>
        ThisWeek = ePivotTableFilterType.ThisWeek,
        /// <summary>
        /// A date filter - This Year
        /// </summary>
        ThisYear = ePivotTableFilterType.ThisYear,
        /// <summary>
        /// A date filter - Today
        /// </summary>
        Today = ePivotTableFilterType.Today,
        /// <summary>
        /// A date filter - Tomorrow
        /// </summary>
        Tomorrow = ePivotTableFilterType.Tomorrow,
        /// <summary>
        /// A date filter - Year to date
        /// </summary>
        YearToDate = ePivotTableFilterType.YearToDate,
        /// <summary>
        /// A date filter - Yesterday
        /// </summary>
        Yesterday = ePivotTableFilterType.Yesterday
    }
    public enum ePivotTableValueFilterType
    {
        /// <summary>
        /// A numeric or string filter - Value Between
        /// </summary>
        ValueBetween = ePivotTableFilterType.ValueBetween,
        /// <summary>
        /// A numeric or string filter - Equal
        /// </summary>
        ValueEqual = ePivotTableFilterType.ValueEqual,
        /// <summary>
        /// A numeric or string filter - GreaterThan
        /// </summary>
        ValueGreaterThan = ePivotTableFilterType.ValueGreaterThan,
        /// <summary>
        /// A numeric or string filter - Greater Than Or Equal
        /// </summary>
        ValueGreaterThanOrEqual = ePivotTableFilterType.ValueGreaterThanOrEqual,
        /// <summary>
        /// A numeric or string filter - Less Than 
        /// </summary>
        ValueLessThan = ePivotTableFilterType.ValueLessThan,
        /// <summary>
        /// A numeric or string filter - Less Than Or Equal
        /// </summary>
        ValueLessThanOrEqual = ePivotTableFilterType.ValueLessThanOrEqual,
        /// <summary>
        /// A numeric or string filter - Not Between
        /// </summary>
        ValueNotBetween = ePivotTableFilterType.ValueNotBetween,
        /// <summary>
        /// A numeric or string filter - Not Equal
        /// </summary>
        ValueNotEqual = ePivotTableFilterType.ValueNotEqual
    }
    public enum ePivotTableTop10FilterType
    {
        /// <summary>
        /// A top/bottom filter - Count
        /// </summary>
        Count = ePivotTableFilterType.Count,
        /// <summary>
        /// A top/bottom filter - Sum
        /// </summary>
        Sum = ePivotTableFilterType.Sum,
        /// <summary>
        /// A top/bottom filter - Percent
        /// </summary>
        Percent = ePivotTableFilterType.Percent
    }
}
