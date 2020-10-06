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
    /// Defines a pivot table date value filter type
    /// </summary>
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
}
