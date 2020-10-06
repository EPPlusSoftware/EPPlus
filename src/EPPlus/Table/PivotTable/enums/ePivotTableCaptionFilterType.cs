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
    /// Defines a pivot table caption filter type
    /// </summary>
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
}
