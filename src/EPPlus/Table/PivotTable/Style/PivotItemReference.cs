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
    /// <summary>
    /// A reference to a pivot table value item
    /// </summary>
    public struct PivotItemReference
    {
        /// <summary>
        /// The index of the item in items of the pivot table field
        /// </summary>
        public int Index { get; internal set; }
        /// <summary>
        /// The value of the item
        /// </summary>
        public object Value { get; internal set; }
    }
}