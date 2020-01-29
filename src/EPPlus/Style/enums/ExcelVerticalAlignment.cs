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
    /// Vertical text alignment
    /// </summary>
    public enum ExcelVerticalAlignment
    {
        /// <summary>
        /// Top aligned
        /// </summary>
        Top,
        /// <summary>
        /// Center aligned
        /// </summary>
        Center,
        /// <summary>
        /// Bottom aligned
        /// </summary>
        Bottom,
        /// <summary>
        /// Distributed. Each line of text inside the cell is evenly distributed across the height of the cell
        /// </summary>
        Distributed,
        /// <summary>
        /// Justify. Each line of text inside the cell is evenly distributed across the height of the cell
        /// </summary>
        Justify
    }
}