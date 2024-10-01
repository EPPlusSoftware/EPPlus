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
    /// The scope of the pivot table conditional formatting rule
    /// </summary>
    public enum ConditionScope
    {
        /// <summary>
        /// The conditional formatting is applied to the selected data fields.
        /// </summary>
        Data,
        /// <summary>
        /// The conditional formatting is applied to the selected PivotTable field intersections.
        /// </summary>
        Field,
        /// <summary>
        /// The conditional formatting is applied to the selected data fields.
        /// </summary>
        Selection
    }
}

