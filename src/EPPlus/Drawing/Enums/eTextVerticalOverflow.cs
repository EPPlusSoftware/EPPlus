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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// How text vertical overflows
    /// </summary>
    public enum eTextVerticalOverflow
    {
        /// <summary>
        /// Clip the text and give no indication that there is text that is not visible at the top and bottom.
        /// </summary>
        Clip,
        /// <summary>
        /// Use an ellipse to highlight text that is not visible at the top and bottom.
        /// </summary>
        Ellipsis,
        /// <summary>
        /// Overflow the text.
        /// </summary>
        Overflow
    }
}