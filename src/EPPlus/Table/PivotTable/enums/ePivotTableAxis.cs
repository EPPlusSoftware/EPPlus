/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
namespace OfficeOpenXml.Table.PivotTable
{
    /*************************************************************************************************
      Required Notice: Copyright (C) EPPlus Software AB. 
      This software is licensed under PolyForm Noncommercial License 1.0.0 
      and may only be used for noncommercial purposes 
      https://polyformproject.org/licenses/noncommercial/1.0.0/

      A commercial license to use this software can be purchased at https://epplussoftware.com
     *************************************************************************************************
      Date               Author                       Change
     *************************************************************************************************
      12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
     *************************************************************************************************/
    /// <summary>
    /// Defines the axis for a pivot table
    /// </summary>
    public enum ePivotTableAxis
    {
        /// <summary>
        /// No axis defined
        /// </summary>
        None,
        /// <summary>
        /// Defines the column axis
        /// </summary>
        ColumnAxis,
        /// <summary>
        /// Defines the page axis
        /// </summary>
        PageAxis,
        /// <summary>
        /// Defines the row axis
        /// </summary>
        RowAxis,
        /// <summary>
        /// Defines the values axis
        /// </summary>
        ValuesAxis
    }
}