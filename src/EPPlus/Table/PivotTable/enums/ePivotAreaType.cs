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
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines the pivot area affected by a style
    /// </summary>
    public enum ePivotAreaType
    {
        /// <summary>
        /// Refers to the whole pivot table
        /// </summary>
        All,
        /// <summary>
        /// Refers to a field button
        /// </summary>
        FieldButton,
        /// <summary>
        /// Refers to data in the data area.
        /// </summary>
        Data,
        /// <summary>
        /// Refers to no pivot area
        /// </summary>
        None,
        /// <summary>
        /// Refers to a header or item
        /// </summary>
        Normal,
        /// <summary>
        /// Refers to the blank cells at the top-left(LTR sheets) or bottom-right(RTL sheets) of the pivot table.
        /// </summary>
        Origin,
        /// <summary>
        /// Refers to the blank cells at the top of the pivot table, on its trailing edge. 
        /// </summary>
        TopEnd
    }
}