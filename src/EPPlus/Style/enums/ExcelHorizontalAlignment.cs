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
namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Horizontal text alignment
    /// </summary>
    public enum ExcelHorizontalAlignment
    {
        /// <summary>
        /// General aligned
        /// </summary>
        General,
        /// <summary>
        /// Left aligned
        /// </summary>
        Left,
        /// <summary>
        /// Center aligned
        /// </summary>
        Center,
        /// <summary>
        /// The horizontal alignment is centered across multiple cells
        /// </summary>
        CenterContinuous,
        /// <summary>
        /// Right aligned
        /// </summary>
        Right,
        /// <summary>
        /// The value of the cell should be filled across the entire width of the cell.
        /// </summary>
        Fill,
        /// <summary>
        /// Each word in each line of text inside the cell is evenly distributed across the width of the cell
        /// </summary>
        Distributed,
        /// <summary>
        /// The horizontal alignment is justified to the Left and Right for each row.
        /// </summary>
        Justify
    }
}