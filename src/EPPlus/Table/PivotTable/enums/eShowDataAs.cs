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
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// The data formats for a field in the PivotTable
    /// </summary>
    public enum eShowDataAs
    {
        /// <summary>
        /// The field is shown as the "difference from" a value.
        /// </summary>
        Difference,
        /// <summary>
        /// The field is shown as the index.
        /// ((Cell Value) x (Grand Total of Grand Totals)) / ((Grand Row Total) x (Grand Column Total))
        /// </summary>
        Index, 
        /// <summary>
        /// The field is shown as its normal datatype.
        /// </summary>
        Normal, 
        /// <summary>
        /// The field is show as the percentage of a value
        /// </summary>
        Percent, 
        /// <summary>
        /// The field is shown as the percentage difference from a value.
        /// </summary>
        PercentDiff, 
        /// <summary>
        /// The field is shown as the percentage of the column.
        /// </summary>
        PercentOfCol,
        /// <summary>
        /// The field is shown as the percentage of the row
        /// </summary>
        PercentOfRow, 
        /// <summary>
        /// The field is shown as the percentage of the total
        /// </summary>
        PercentOfTotal, 
        /// <summary>
        /// The field is shown as the running total in the the table
        /// </summary>
        RunTotal,
        /// <summary>
        /// The field is shown as the percentage of the parent row total
        /// </summary>
        PercentOfParentRow,
        /// <summary>
        /// The field is shown as the percentage of the parent column total
        /// </summary>
        PercentOfParentCol,
        /// <summary>
        /// The field is shown as the percentage of the parent total
        /// </summary>
        PercentOfParent,
        /// <summary>
        /// The field is shown as the rank ascending.
        /// Lists the smallest item in the field as 1, and each larger value with a higher rank value.
        /// </summary>
        RankAscending,
        /// <summary>
        /// The field is shown as the rank descending.
        /// Lists the largest item in the field as 1, and each smaller value with a higher rank value.
        /// </summary>
        RankDescending,
        /// <summary>
        /// The field is shown as the percentage of the running total
        /// </summary>
        PercentOfRunningTotal

    }
}
