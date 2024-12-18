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
    /// Conditional Formatting Evaluation Type
    /// </summary>
    public enum ConditionType
    {
        /// <summary>
        /// The conditional formatting is not evaluated
        /// </summary>
        None,
        /// <summary>
        /// The Top N conditional formatting is evaluated across the entire scope range.
        /// </summary>
        All,
        /// <summary>
        /// The Top N conditional formatting is evaluated for each row§.
        /// </summary>
        Row,
        /// <summary>
        /// The Top N conditional formatting is evaluated for each column.
        /// </summary>
        Column
    }
}

