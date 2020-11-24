/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// The reading order
    /// </summary>
    public enum eReadingOrder
    {
        /// <summary>
        /// Reading order is determined by the first non-whitespace character
        /// </summary>
        ContextDependent = 0,
        /// <summary>
        /// Left to Right
        /// </summary>
        LeftToRight = 1,
        /// <summary>
        /// Right to Left
        /// </summary>
        RightToLeft = 2
    }
}