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
    /// Defines a pivot table top 10 filter type
    /// </summary>
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
