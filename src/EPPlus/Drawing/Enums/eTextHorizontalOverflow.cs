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
    /// Specifies the text vertical overflow
    /// </summary>
    public enum eTextHorizontalOverflow
    {
        /// <summary>
        /// When a character doesn't fit into a line, clip it at the end.
        /// </summary>
        Clip,
        /// <summary>
        /// When a character doesn't fit into a line, allow an overflow.
        /// </summary>
        Overflow
    }
}