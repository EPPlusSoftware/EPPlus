/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date (MM/DD/YYYY)              Author                       Change
 *************************************************************************************************
  06/10/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// In which way to store the position of a chart element
    /// </summary>
    public enum eLayoutMode
    {
        /// <summary>
        /// Specifies that the Width or Height shall be interpreted
        /// as the Right or Bottom of the chart element.
        /// </summary>
        Edge,
        /// <summary>
        /// Specifies that the Width or Height shall be interpreted
        /// as the Width or Height of the chart element.
        /// </summary>
        Factor
    }
}
