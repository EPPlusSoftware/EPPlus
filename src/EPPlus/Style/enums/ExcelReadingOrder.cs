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
    /// The reading order
    /// </summary>
    public enum ExcelReadingOrder
    {
        /// <summary>
        /// Reading order is determined by the first non-whitespace character
        /// </summary>
        ContextDependent=0,
        /// <summary>
        /// Left to Right
        /// </summary>
        LeftToRight=1,
        /// <summary>
        /// Right to Left
        /// </summary>
        RightToLeft=2
    }
}